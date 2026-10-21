<#
.SYNOPSIS
  OSign - the central editor for the Outlook signatures the add-in inserts.

.DESCRIPTION
  Starts a small web server on 127.0.0.1, opens the editor in a Chrome or Edge app window, and
  publishes the signatures to where the add-in reads them: the OSign folder in the company's
  SharePoint, synced to this PC by OneDrive (or any other folder). Nothing goes anywhere else.

  This script is one backend for the web UI in .\ui. The UI talks to it only through the HTTP API
  described in docs/API.md, so another backend - a container - can take its place without the UI
  changing.

  Data: %APPDATA%\OSign (or $env:OSIGN_DATA): the draft, its images, the settings.

.PARAMETER NoBrowser
  Start the server without opening a window. The address to open is printed.

.PARAMETER Port
  Listen on this port instead of the first free one of 17690-17697.

.PARAMETER ExportUsers
  No window, no server: export the users' values from Active Directory into the publish folder
  (users\<mail>.json), as the published signatures.json asks for, and exit. For the Task
  Scheduler, so the values in signatures stay as current as AD:
    powershell -ExecutionPolicy Bypass -File OSign.ps1 -ExportUsers
#>
param(
    [switch]$NoBrowser,
    [int]$Port = 0,
    [switch]$ExportUsers
)

$ErrorActionPreference = 'Stop'
$script:AppVersion = '0.4.0'
$script:OnWindows = ($PSVersionTable.PSEdition -eq 'Desktop') -or ($IsWindows -eq $true)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

# ---------------------------------------------------------------------------
# Where things are
# ---------------------------------------------------------------------------
$script:UiDir = Join-Path $PSScriptRoot 'ui'
# The engine shared with the add-in: next to the script in a release, one folder up in the repo.
$script:CoreFile = @((Join-Path $PSScriptRoot 'shared\signature-core.js'), (Join-Path $PSScriptRoot '..\shared\signature-core.js')) |
    ForEach-Object { [IO.Path]::GetFullPath(($_ -replace '\\', [IO.Path]::DirectorySeparatorChar)) } |
    Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
# The add-in's files, with blanks for its settings (addin/build.mjs): OSign\addin in a release,
# dist\template in the repo after npm run build.
$script:AddinTemplateDir = @((Join-Path $PSScriptRoot 'addin'), (Join-Path $PSScriptRoot '..\dist\template')) |
    ForEach-Object { [IO.Path]::GetFullPath(($_ -replace '\\', [IO.Path]::DirectorySeparatorChar)) } |
    Where-Object { Test-Path -LiteralPath (Join-Path $_ 'osign-addin.json') } | Select-Object -First 1

function Get-DataDir {
    if ($env:OSIGN_DATA) { return $env:OSIGN_DATA }
    if ($env:APPDATA) { return (Join-Path $env:APPDATA 'OSign') }
    return (Join-Path $HOME '.osign')
}
$script:DataDir = Get-DataDir
$script:DraftFile = Join-Path $script:DataDir 'draft.json'
$script:AssetDir = Join-Path $script:DataDir 'assets'
$script:SettingsFile = Join-Path $script:DataDir 'settings.json'
$script:StateFile = Join-Path $script:DataDir 'state.json'
foreach ($d in @($script:DataDir, $script:AssetDir)) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:MaxAssetBytes = 1MB
$script:AssetTypes = @{ '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.gif' = 'image/gif' }

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
function Read-Text([string]$Path) { if (Test-Path -LiteralPath $Path) { return [IO.File]::ReadAllText($Path, $script:Utf8) } return $null }
# Written to a temporary file and moved over the old one, so a crash never leaves half a file.
function Write-Text([string]$Path, [string]$Text) {
    $tmp = "$Path.tmp"
    [IO.File]::WriteAllText($tmp, $Text, $script:Utf8)
    if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($tmp, $Path, [NullString]::Value) } else { [IO.File]::Move($tmp, $Path) }
}
function To-Json($obj) { return (ConvertTo-Json -InputObject $obj -Depth 20 -Compress) }
function Ok([hashtable]$h) { if (-not $h) { $h = @{} }; $h.ok = $true; return (To-Json $h) }
function Fail([string]$msg) { return (To-Json @{ ok = $false; error = $msg }) }
function Get-Sha256Hex([byte[]]$Bytes) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return (($sha.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $sha.Dispose() }
}

# An asset's file name: what the upload was called, cut down to what is safe in a path and a URL.
# Returns $null for anything that is not one of the image types Outlook shows.
function Get-SafeAssetName([string]$Name) {
    # Only the last part of a path, whichever slash it uses.
    $base = ([string]$Name -split '[\\/]')[-1]
    $ext = [IO.Path]::GetExtension($base).ToLowerInvariant()
    if (-not $script:AssetTypes.ContainsKey($ext)) { return $null }
    $stem = [IO.Path]::GetFileNameWithoutExtension($base) -replace '[^A-Za-z0-9_-]+', '-'
    $stem = $stem.Trim('-')
    if (-not $stem) { $stem = 'image' }
    if ($stem.Length -gt 60) { $stem = $stem.Substring(0, 60) }
    return ($stem + $ext)
}

# The LDAP filter syntax's special characters, escaped (RFC 4515), so a lookup for "a*)(x" is that text.
function ConvertTo-LdapFilterValue([string]$Value) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '\' { [void]$sb.Append('\5c') }
            '*' { [void]$sb.Append('\2a') }
            '(' { [void]$sb.Append('\28') }
            ')' { [void]$sb.Append('\29') }
            ([char]0) { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
# folder:       where Publish writes - the OSign folder of SharePoint, synced by OneDrive
#               (e.g. C:\Users\me\Contoso\Contoso - Documents\OSign). The add-in reads it there.
# photosFolder: where the preview finds the photos per person (<username>.jpg). Empty: the
#               photos folder inside the publish folder.
# The Outlook add-in's files (Build-AddinPackage):
# addinUrl:     the HTTPS address they are served from, e.g. https://osign.contoso.com/
# clientId:     the Entra ID app registration's Application (client) ID
# tenantId:     only for an app registration of a single organization
# spSite:       the SharePoint site holding the publish folder: root (the home site) or
#               <host>.sharepoint.com:/sites/<name>: (or /teams/<name>), or someone's OneDrive:
#               <host>-my.sharepoint.com:/personal/<name>:
# spFolder:     that folder in the site's document library, e.g. OSign
# addinFolder:  where the finished files are written, to be copied to the web server
# addinHost:    the kind of web server (cloudflare, azure, iis, apache, other): its caching file
#               goes with the files
# setupDone:    '1' once the setup assistant was finished; until then it opens by itself
function Get-DefaultSettings { return [ordered]@{ folder = ''; photosFolder = ''; addinUrl = ''; clientId = ''; tenantId = ''; spSite = 'root'; spFolder = 'OSign'; addinFolder = ''; addinHost = ''; setupDone = '' } }

function Get-Settings {
    $d = Get-DefaultSettings
    $raw = Read-Text $script:SettingsFile
    if ($raw) {
        try {
            $s = $raw | ConvertFrom-Json
            foreach ($k in @($d.Keys)) { if ($null -ne $s.$k) { $d[$k] = [string]$s.$k } }
        } catch { }
    }
    return $d
}

function Test-Settings($s, [switch]$ForPublish) {
    foreach ($k in @('folder', 'photosFolder', 'addinFolder')) {
        $v = [string]$s[$k]
        if ($v -and ($v.IndexOfAny([IO.Path]::GetInvalidPathChars()) -ge 0 -or -not [IO.Path]::IsPathRooted($v))) { return "The $k is a full path, like C:\Users\me\Contoso\Documents\OSign." }
    }
    $err = Test-AddinSettings $s; if ($err) { return $err }
    if ($ForPublish) {
        if (-not $s.folder) { return 'Pick the folder to publish to in Settings: the OSign folder of SharePoint, synced to this PC.' }
        if (-not (Test-Path -LiteralPath $s.folder -PathType Container)) { return "The folder $($s.folder) is not there. Is the SharePoint library synced to this PC?" }
    }
    return $null
}

# The same rules as checkSetup in addin/build.mjs. Empty values pass; Build-AddinPackage needs the
# address.
$script:GuidRx = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
function Test-AddinSettings($s) {
    $u = [string]$s.addinUrl
    if ($u -and $u -notmatch '^https://[^\s"<>''&]+$') { return 'The add-in address is an https:// address, like https://osign.contoso.com/.' }
    if ($s.clientId -and $s.clientId -notmatch $script:GuidRx) { return "The client ID is the app registration's Application (client) ID, a GUID." }
    if ($s.tenantId -and $s.tenantId -notmatch $script:GuidRx) { return 'The tenant ID is a Directory (tenant) ID, a GUID.' }
    if ($s.spSite -and $s.spSite -notmatch '^root$|^[A-Za-z0-9.-]+\.sharepoint\.com:/(sites|teams|personal)/[^:/"\\\s]+:$') { return 'The SharePoint site is root, <host>.sharepoint.com:/sites/<name>: or, for a OneDrive, <host>-my.sharepoint.com:/personal/<name>:' }
    $f = ([string]$s.spFolder).Trim('/')
    if ($s.addinHost -and @('cloudflare', 'azure', 'iis', 'apache', 'other') -notcontains $s.addinHost) { return 'The web server is cloudflare, azure, iis, apache or other.' }
    if ($s.spFolder -and ($f -match '["\\\x00-\x1f]' -or @($f -split '/' | Where-Object { -not $_ -or $_ -eq '.' -or $_ -eq '..' }).Count)) { return 'The SharePoint folder is a folder like OSign.' }
    return $null
}

function Get-PhotosDir {
    $s = Get-Settings
    if ($s.photosFolder) { return $s.photosFolder }
    if ($s.folder) { return (Join-Path $s.folder 'photos') }
    return $null
}

# ---------------------------------------------------------------------------
# Draft: the configuration being edited, and its images
# ---------------------------------------------------------------------------
function Get-DraftAssets {
    $list = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:AssetDir -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        if (Get-SafeAssetName $f.Name) { $list += [ordered]@{ name = $f.Name; size = $f.Length } }
    }
    return , $list
}

function Get-State {
    $raw = Read-Text $script:StateFile
    if ($raw) { try { return ($raw | ConvertFrom-Json) } catch { } }
    return [pscustomobject]@{ publishedHash = ''; publishedAt = 0 }
}
function Save-State($hash) {
    # A number, not a date text: ConvertFrom-Json in PowerShell 7 turns ISO dates into DateTime,
    # which then comes back in the machine's own date format.
    Write-Text $script:StateFile (To-Json ([ordered]@{ publishedHash = $hash; publishedAt = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() }))
}

# One hash over the configuration and every image, to tell whether the draft differs from what was
# last published or loaded.
function Get-DraftHash {
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add((Get-Sha256Hex ($script:Utf8.GetBytes([string](Read-Text $script:DraftFile)))))
    foreach ($a in (Get-DraftAssets)) { $parts.Add($a.name + ':' + (Get-Sha256Hex ([IO.File]::ReadAllBytes((Join-Path $script:AssetDir $a.name))))) }
    return (Get-Sha256Hex ($script:Utf8.GetBytes(($parts -join "`n"))))
}

# The UI sends the configuration as JSON text, and it is stored as that text: going through
# PowerShell objects and back would change it (lists of one, empty lists, number formats).
function Test-ConfigJson([string]$Json) {
    if (-not $Json) { return 'No configuration was sent.' }
    # Windows PowerShell 5.1 reads no JSON over 2 MB (ConvertFrom-Json), so a request has to stay below.
    if ($Json.Length -gt 1500KB) { return 'The configuration is larger than 1.5 MB: images belong in assets, not in the HTML.' }
    try { $c = $Json | ConvertFrom-Json } catch { return 'The configuration is not valid JSON.' }
    if ($c -isnot [pscustomobject]) { return 'The configuration is not a JSON object.' }
    if ($c.schemaVersion -ne 1) { return 'The configuration has an unknown schemaVersion.' }
    return $null
}

# ---------------------------------------------------------------------------
# Publishing: the folder gets signatures.json and assets\<image>. Its photos\ are not touched:
# they are put there by hand, once (one <username>.jpg per person).
# ---------------------------------------------------------------------------
function Publish-Folder([string]$Root, [string]$ConfigJson) {
    $assets = Join-Path $Root 'assets'
    if (-not (Test-Path -LiteralPath $assets)) { New-Item -ItemType Directory -Path $assets -Force | Out-Null }
    $keep = @{}; $changed = $false
    foreach ($a in (Get-DraftAssets)) {
        $keep[$a.name] = 1
        $src = Join-Path $script:AssetDir $a.name; $dst = Join-Path $assets $a.name
        # Only what differs is written: OneDrive then uploads nothing that did not change.
        if (-not (Test-Path -LiteralPath $dst) -or (Get-Sha256Hex ([IO.File]::ReadAllBytes($dst))) -ne (Get-Sha256Hex ([IO.File]::ReadAllBytes($src)))) {
            Copy-Item -LiteralPath $src -Destination $dst -Force; $changed = $true
        }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $assets -File)) {
        if (-not $keep.ContainsKey($f.Name) -and (Get-SafeAssetName $f.Name)) { Remove-Item -LiteralPath $f.FullName -Force; $changed = $true }
    }
    # The configuration last: the add-in never reads one that points at images not there yet.
    $cfgPath = Join-Path $Root 'signatures.json'
    if ((Read-Text $cfgPath) -ne $ConfigJson) { Write-Text $cfgPath $ConfigJson; $changed = $true }
    return @{ changed = $changed }
}

function Read-FolderPublished([string]$Root) {
    $out = @{ config = (Read-Text (Join-Path $Root 'signatures.json')); assets = @{} }
    foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $Root 'assets') -File -ErrorAction SilentlyContinue)) {
        $n = Get-SafeAssetName $f.Name; if ($n -eq $f.Name) { $out.assets[$n] = [IO.File]::ReadAllBytes($f.FullName) }
    }
    return $out
}

# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Where a synced folder is in SharePoint: OneDrive notes, for each library it syncs to this PC,
# the local folder (MountPoint) and the library's address (UrlNamespace). From the longest one the
# path starts with, the folder's address in SharePoint follows, and from that what the add-in needs:
# the site (as Graph names it) and the folder within the site's default library.
# ---------------------------------------------------------------------------
function Get-SyncedLibraries {
    if ($env:OSIGN_TEST_SYNC_JSON) {
        $list = (Read-Text $env:OSIGN_TEST_SYNC_JSON) | ConvertFrom-Json
        return , @(@($list) | ForEach-Object { [pscustomobject]@{ MountPoint = [string]$_.MountPoint; UrlNamespace = [string]$_.UrlNamespace } })
    }
    $out = @()
    if (-not $script:OnWindows) { return , $out }
    foreach ($k in @(Get-ChildItem -Path 'HKCU:\Software\SyncEngines\Providers\OneDrive' -ErrorAction SilentlyContinue)) {
        try {
            $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction Stop
            if ($p.MountPoint -and $p.UrlNamespace) { $out += [pscustomobject]@{ MountPoint = [string]$p.MountPoint; UrlNamespace = [string]$p.UrlNamespace } }
        } catch { }
    }
    return , $out
}

# The names the default library's address has: "Shared Documents" on a site, "Documents" in a
# OneDrive (also in German tenants; older German sites may have "Freigegebene Dokumente").
$script:DefaultLibraries = @('shared documents', 'documents', 'freigegebene dokumente', 'dokumente')

function Resolve-SpLocation([string]$Path) {
    if (-not $Path) { return $null }
    $norm = { param($x) ([string]$x).Replace('/', '\').TrimEnd('\') }
    $full = & $norm $Path
    $best = $null
    foreach ($l in (Get-SyncedLibraries)) {
        $mp = & $norm $l.MountPoint
        if (-not $mp) { continue }
        if ($full.Equals($mp, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($mp + '\', [StringComparison]::OrdinalIgnoreCase)) {
            if (-not $best -or $mp.Length -gt (& $norm $best.MountPoint).Length) { $best = $l }
        }
    }
    if (-not $best) { return $null }
    $rel = $full.Substring((& $norm $best.MountPoint).Length).Trim('\')
    $url = $best.UrlNamespace.TrimEnd('/') + '/' + ($rel -replace '\\', '/')
    $u = [Uri]$url.TrimEnd('/')
    $parts = @($u.AbsolutePath.Split('/') | Where-Object { $_ } | ForEach-Object { [Uri]::UnescapeDataString($_) })
    $site = 'root'; $siteName = ''
    if ($parts.Count -ge 2 -and $parts[0] -match '^(sites|teams|personal)$') {
        $site = $u.Host.ToLowerInvariant() + ':/' + $parts[0].ToLowerInvariant() + '/' + $parts[1] + ':'
        $siteName = $parts[1]
        $parts = @($parts | Select-Object -Skip 2)
    }
    $library = $(if ($parts.Count) { $parts[0] } else { '' })
    $folder = (@($parts | Select-Object -Skip 1) -join '/')
    return [ordered]@{
        site = $site; siteName = $siteName; host = $u.Host.ToLowerInvariant(); library = $library; folder = $folder
        url = $url.TrimEnd('/'); defaultLibrary = ($script:DefaultLibraries -contains $library.ToLowerInvariant())
    }
}

function Api-SpResolve($data) {
    $r = Resolve-SpLocation ([string]$data.path)
    if (-not $r) { return (Ok @{ found = $false }) }
    return (Ok @{ found = $true; location = $r })
}

# The Windows folder dialog, for the folder fields. In a Windows PowerShell of its own: the dialog
# needs a single-threaded apartment, which PowerShell 7 does not have. The path comes back through
# a file, so names with umlauts arrive as they are.
function Api-PickFolder($data) {
    if (-not $script:OnWindows) { return (Fail 'Choosing a folder needs Windows.') }
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $start = ([string]$data.start).Replace("'", "''")
        $title = ([string]$data.title).Replace("'", "''")
        $outFile = $tmp.Replace("'", "''")
        $code = @"
Add-Type -AssemblyName System.Windows.Forms
`$owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = `$true; ShowInTaskbar = `$false }
`$d = New-Object System.Windows.Forms.FolderBrowserDialog
`$d.Description = '$title'
`$d.ShowNewFolderButton = `$true
if ('$start' -and (Test-Path -LiteralPath '$start')) { `$d.SelectedPath = '$start' }
if (`$d.ShowDialog(`$owner) -eq [System.Windows.Forms.DialogResult]::OK) { [IO.File]::WriteAllText('$outFile', `$d.SelectedPath, (New-Object System.Text.UTF8Encoding(`$false))) }
"@
        $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
        & $ps -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -EncodedCommand $enc | Out-Null
        $picked = [IO.File]::ReadAllText($tmp, $script:Utf8)
        return (Ok @{ path = $picked })
    } finally { Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue }
}

function Api-Info {
    $ad = $false
    if ($script:OnWindows) {
        try { Add-Type -AssemblyName System.DirectoryServices -ErrorAction SilentlyContinue; $ad = [bool]([DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()) } catch { $ad = $false }
    }
    if ($env:OSIGN_TEST_AD_JSON) { $ad = $true }
    return (Ok @{ backend = 'powershell'; version = $script:AppVersion; dataDir = $script:DataDir; features = @{ adLookup = $ad; addinFiles = [bool]$script:AddinTemplateDir; pickFolder = [bool]$script:OnWindows } })
}

function Api-DraftGet {
    $st = Get-State
    $cfg = Read-Text $script:DraftFile
    return (Ok @{
        configJson = $cfg; assets = (Get-DraftAssets)
        dirty = ([bool]$cfg -and ((Get-DraftHash) -ne [string]$st.publishedHash))
        publishedAt = $(if ($st.publishedAt -is [string]) { 0 } else { [int64]$st.publishedAt })
    })
}

function Api-DraftSave($data) {
    $err = Test-ConfigJson ([string]$data.configJson); if ($err) { return (Fail $err) }
    Write-Text $script:DraftFile ([string]$data.configJson)
    return (Ok @{ dirty = ((Get-DraftHash) -ne [string](Get-State).publishedHash) })
}

function Api-AssetUpload($data) {
    $name = Get-SafeAssetName ([string]$data.name)
    if (-not $name) { return (Fail 'Only PNG, JPG and GIF images can be used: Outlook does not show SVG or WebP in a signature.') }
    try { $bytes = [Convert]::FromBase64String([string]$data.data) } catch { return (Fail 'The image did not arrive whole.') }
    if ($bytes.Length -eq 0) { return (Fail 'The image is empty.') }
    if ($bytes.Length -gt $script:MaxAssetBytes) { return (Fail 'The image is larger than 1 MB. A signature image is best well under 100 KB.') }
    # The content has to be the image it says it is.
    $sig = ($bytes[0..([Math]::Min(7, $bytes.Length - 1))] | ForEach-Object { $_.ToString('x2') }) -join ''
    $ext = [IO.Path]::GetExtension($name)
    $okType = ($ext -eq '.png' -and $sig.StartsWith('89504e47')) -or ($ext -in @('.jpg', '.jpeg') -and $sig.StartsWith('ffd8ff')) -or ($ext -eq '.gif' -and $sig.StartsWith('47494638'))
    if (-not $okType) { return (Fail "The file is not the $($ext.TrimStart('.').ToUpper()) image its name says.") }
    # Same name, other content: a new name, so a signature already using the old image keeps it.
    $path = Join-Path $script:AssetDir $name
    if ((Test-Path -LiteralPath $path) -and ((Get-Sha256Hex ([IO.File]::ReadAllBytes($path))) -ne (Get-Sha256Hex $bytes))) {
        $stem = [IO.Path]::GetFileNameWithoutExtension($name); $i = 2
        do { $name = "$stem-$i$ext"; $path = Join-Path $script:AssetDir $name; $i++ } while (Test-Path -LiteralPath $path)
    }
    [IO.File]::WriteAllBytes($path, $bytes)
    return (Ok @{ path = "assets/$name"; name = $name; size = $bytes.Length })
}

function Api-AssetDelete($data) {
    $name = Get-SafeAssetName ([string]$data.name)
    if (-not $name -or $name -ne [string]$data.name) { return (Fail 'No such image.') }
    Remove-Item -LiteralPath (Join-Path $script:AssetDir $name) -Force -ErrorAction SilentlyContinue
    return (Ok @{})
}

function Api-SettingsGet {
    $s = Get-Settings
    $pd = Get-PhotosDir
    return (Ok @{ settings = $s; folderFound = [bool]($s.folder -and (Test-Path -LiteralPath $s.folder -PathType Container)); photosFound = [bool]($pd -and (Test-Path -LiteralPath $pd -PathType Container)) })
}

# Only what is sent changes; the rest stays as it was.
function Api-SettingsSave($data) {
    $s = Get-Settings
    foreach ($k in @($s.Keys)) { if ($null -ne $data.settings.$k) { $s[$k] = ([string]$data.settings.$k).Trim().TrimEnd('\', '/') } }
    $err = Test-Settings $s; if ($err) { return (Fail $err) }
    Write-Text $script:SettingsFile (To-Json $s)
    return (Api-SettingsGet)
}

# ---------------------------------------------------------------------------
# The Outlook add-in's files: the template filled in with the settings, for any HTTPS web server
# ---------------------------------------------------------------------------
function ConvertTo-XmlText([string]$Text) { return [Security.SecurityElement]::Escape($Text) }

function Build-AddinPackage($s) {
    if (-not $script:AddinTemplateDir) { return @{ error = 'The add-in files are not next to OSign.ps1 (folder addin). Use the release zip, or run npm run build in the repository.' } }
    $err = Test-AddinSettings $s; if ($err) { return @{ error = $err } }
    if (-not $s.addinUrl) { return @{ error = 'Enter the address the add-in will be served from.' } }
    if (-not $s.clientId) { return @{ error = 'Enter the app registration''s client ID: without it the add-in cannot sign in.' } }
    if (-not $s.addinFolder) { return @{ error = 'Enter the folder to write the add-in''s files to.' } }
    $out = [IO.Path]::GetFullPath($s.addinFolder)
    foreach ($other in @($s.folder, $script:DataDir)) {
        if ($other -and ($out.TrimEnd('\', '/') -eq [IO.Path]::GetFullPath($other).TrimEnd('\', '/'))) { return @{ error = 'Write the add-in''s files to a folder of their own, not the publish folder.' } }
    }
    $base = $s.addinUrl.TrimEnd('/') + '/'
    $origin = ([Uri]$base).GetLeftPart([UriPartial]::Authority)
    $js = [ordered]@{ '@@OSIGN_CLIENT_ID@@' = $s.clientId; '@@OSIGN_TENANT_ID@@' = $s.tenantId; '@@OSIGN_SP_SITE@@' = $(if ($s.spSite) { $s.spSite } else { 'root' }); '@@OSIGN_SP_FOLDER@@' = $(if ($s.spFolder) { $s.spFolder.Trim('/') } else { 'OSign' }) }
    $xml = [ordered]@{ '{{BASE}}' = (ConvertTo-XmlText $base); '{{ORIGIN}}' = (ConvertTo-XmlText $origin); '{{SUPPORT}}' = (ConvertTo-XmlText $base) }
    $fill = @{ 'launchevent.js' = $js; 'taskpane.js' = $js; 'manifest.xml' = $xml }
    $root = $script:AddinTemplateDir.TrimEnd('\', '/')
    $n = 0
    $sep = [IO.Path]::DirectorySeparatorChar
    $hostDir = $(if ($s.addinHost) { 'hosting' + $sep + $s.addinHost + $sep } else { $null })
    foreach ($f in @(Get-ChildItem -LiteralPath $root -File -Recurse -Force)) {
        $rel = $f.FullName.Substring($root.Length + 1)
        # hosting\<host>\ holds each kind of web server's caching file: only the one chosen goes,
        # next to the others.
        if ($rel.StartsWith('hosting' + $sep)) {
            if (-not $hostDir -or -not $rel.StartsWith($hostDir)) { continue }
            $rel = $rel.Substring($hostDir.Length)
        }
        $dest = Join-Path $out $rel
        $dir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ($fill.ContainsKey($rel)) {
            $text = [IO.File]::ReadAllText($f.FullName, $script:Utf8)
            foreach ($k in $fill[$rel].Keys) { $text = $text.Replace($k, [string]$fill[$rel][$k]) }
            [IO.File]::WriteAllText($dest, $text, $script:Utf8)
        } else {
            [IO.File]::Copy($f.FullName, $dest, $true)
        }
        $n++
    }
    $version = ([string](Read-Text (Join-Path $root 'osign-addin.json')) | ConvertFrom-Json).version
    return @{ folder = $out; files = $n; version = $version; manifestUrl = $base + 'manifest.xml'; redirectUri = 'brk-multihub://' + ([Uri]$base).Authority }
}

# Whether the address serves the files made here: the manifest and both bundles as they are in the
# add-in's folder, and the caching header. Answers a list of { what, ok, detail }.
function Test-AddinOnline([string]$Base, [string]$Dir) {
    $Base = $Base.TrimEnd('/') + '/'
    $checks = @()
    foreach ($name in @('manifest.xml', 'launchevent.js', 'taskpane.js', 'taskpane.html')) {
        $local = Join-Path $Dir $name
        try {
            $r = Invoke-WebRequest -Uri ($Base + $name) -UseBasicParsing -TimeoutSec 20 -MaximumRedirection 5 -Headers @{ 'Cache-Control' = 'no-cache' }
            # Byte for byte: the text as decoded would depend on the charset the server names.
            $same = (Test-Path -LiteralPath $local) -and ((Get-Sha256Hex $r.RawContentStream.ToArray()) -eq (Get-Sha256Hex ([IO.File]::ReadAllBytes($local))))
            $checks += [ordered]@{ what = $name; ok = $same; detail = $(if ($same) { 'the same as made here' } else { 'there, but not the files made here: upload them again' }) }
            if ($name -eq 'manifest.xml') {
                $cc = [string]$r.Headers['Cache-Control']
                $checks += [ordered]@{ what = 'Cache-Control'; ok = ($cc -match 'no-cache|max-age=0'); detail = $(if ($cc) { $cc } else { 'not sent: updates can take a while to reach everyone' }) }
            }
        } catch {
            $checks += [ordered]@{ what = $name; ok = $false; detail = 'cannot be loaded: ' + $_.Exception.Message }
        }
    }
    return , $checks
}

function Api-AddinCheck {
    $s = Get-Settings
    if (-not $s.addinUrl) { return (Fail 'Enter the address the add-in will be served from.') }
    if (-not $s.addinFolder -or -not (Test-Path -LiteralPath (Join-Path $s.addinFolder 'manifest.xml'))) { return (Fail 'Make the add-in files first.') }
    $checks = Test-AddinOnline $s.addinUrl $s.addinFolder
    return (Ok @{ checks = $checks; allOk = -not @($checks | Where-Object { -not $_.ok -and $_.what -ne 'Cache-Control' }).Count })
}

function Api-AddinBuild {
    $r = Build-AddinPackage (Get-Settings)
    if ($r.error) { return (Fail $r.error) }
    return (Ok $r)
}

function Api-Publish($data) {
    $err = Test-ConfigJson ([string]$data.configJson); if ($err) { return (Fail $err) }
    $s = Get-Settings
    $err = Test-Settings $s -ForPublish; if ($err) { return (Fail $err) }
    Write-Text $script:DraftFile ([string]$data.configJson)
    try { $r = Publish-Folder $s.folder ([string]$data.configJson) } catch { return (Fail "Writing to $($s.folder): $($_.Exception.Message)") }
    Save-State (Get-DraftHash)
    return (Ok @{ changed = $r.changed; folder = $s.folder })
}

function Api-LoadPublished {
    $s = Get-Settings
    $err = Test-Settings $s -ForPublish; if ($err) { return (Fail $err) }
    try { $p = Read-FolderPublished $s.folder } catch { return (Fail $_.Exception.Message) }
    if (-not $p.config) { return (Fail 'Nothing has been published there yet.') }
    $err = Test-ConfigJson $p.config; if ($err) { return (Fail "The published configuration: $err") }
    # The draft being replaced is kept, once, in case it held work that was never published.
    if (Test-Path -LiteralPath $script:DraftFile) { Copy-Item -LiteralPath $script:DraftFile -Destination "$script:DraftFile.bak" -Force }
    Write-Text $script:DraftFile $p.config
    foreach ($f in @(Get-ChildItem -LiteralPath $script:AssetDir -File)) { if (-not $p.assets.ContainsKey($f.Name)) { Remove-Item -LiteralPath $f.FullName -Force } }
    foreach ($n in $p.assets.Keys) { [IO.File]::WriteAllBytes((Join-Path $script:AssetDir $n), $p.assets[$n]) }
    Save-State (Get-DraftHash)
    return (Api-DraftGet)
}

# The values a user's signature would get, read from Active Directory, for the preview.
function Api-AdLookup($data) {
    if (-not $script:OnWindows) { return (Fail 'Looking up users needs Windows on a machine in the domain.') }
    $q = ([string]$data.query).Trim(); if (-not $q) { $q = $env:USERNAME }
    $attrs = @()
    foreach ($f in @($data.fields)) { $a = [string]$f.ad; if ($a -match '^[A-Za-z][A-Za-z0-9-]{0,63}$' -and $attrs -notcontains $a) { $attrs += $a } }
    $v = ConvertTo-LdapFilterValue $q
    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction SilentlyContinue
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$v)(mail=$v)(userPrincipalName=$v)))"
        $searcher.SizeLimit = 2
        foreach ($a in (@('displayName', 'sAMAccountName', 'distinguishedName') + $attrs)) { [void]$searcher.PropertiesToLoad.Add($a) }
        $found = @($searcher.FindAll())
    } catch { return (Fail "Active Directory: $($_.Exception.Message)") }
    if ($found.Count -eq 0) { return (Fail "No user $q in Active Directory.") }
    if ($found.Count -gt 1) { return (Fail "More than one user matches $q.") }
    $props = $found[0].Properties
    $values = @{}
    foreach ($f in @($data.fields)) {
        $a = [string]$f.ad; $k = [string]$f.key
        if (-not $k) { continue }
        $values[$k] = $(if ($a -and $props[$a.ToLowerInvariant()].Count) { [string]$props[$a.ToLowerInvariant()][0] } else { '' })
    }
    $name = $(if ($props['displayname'].Count) { [string]$props['displayname'][0] } else { $q })
    # Every group the user is in, nested ones included (the add-in reads them the same way from
    # Entra ID): the matching rule IN_CHAIN follows group in group.
    $groups = @()
    try {
        $dn = ConvertTo-LdapFilterValue ([string]$props['distinguishedname'][0])
        $gs = New-Object System.DirectoryServices.DirectorySearcher
        $gs.Filter = "(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=$dn))"
        $gs.PageSize = 500
        foreach ($a in @('name', 'sAMAccountName')) { [void]$gs.PropertiesToLoad.Add($a) }
        foreach ($g in @($gs.FindAll())) {
            foreach ($a in @('samaccountname', 'name')) { if ($g.Properties[$a].Count) { $n = [string]$g.Properties[$a][0]; if ($groups -notcontains $n) { $groups += $n } } }
        }
    } catch { }
    return (Ok @{ values = $values; displayName = $name; groups = @($groups | Sort-Object) })
}

# Groups whose name starts with what was typed, for picking who gets a signature.
function Api-AdGroups($data) {
    if (-not $script:OnWindows) { return (Fail 'Looking up groups needs Windows on a machine in the domain.') }
    $q = ([string]$data.query).Trim()
    if ($q.Length -lt 2) { return (Ok @{ groups = @() }) }
    $v = ConvertTo-LdapFilterValue $q
    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction SilentlyContinue
        $gs = New-Object System.DirectoryServices.DirectorySearcher
        $gs.Filter = "(&(objectClass=group)(|(sAMAccountName=$v*)(name=$v*)))"
        $gs.SizeLimit = 30
        [void]$gs.PropertiesToLoad.Add('sAMAccountName')
        $names = @(@($gs.FindAll()) | ForEach-Object { [string]$_.Properties['samaccountname'][0] } | Where-Object { $_ } | Sort-Object -Unique)
    } catch { return (Fail "Active Directory: $($_.Exception.Message)") }
    return (Ok @{ groups = $names })
}

# ---------------------------------------------------------------------------
# Users from AD, for signatures whose values come from AD (options.valueSource 'ad' or 'mixed'):
# one file per user, users\<mail>.json in the publish folder, with the fields' values and the
# groups (of those the signatures name) the user is in, nested ones included. The add-in reads
# only the file of the user who writes.
# ---------------------------------------------------------------------------
# The fields to export, as signature-core.js's fieldSource() decides it.
function Get-AdExportFields($Config) {
    $mode = [string]$Config.options.valueSource
    $out = @()
    foreach ($f in @($Config.fields)) {
        $ad = [string]$f.ad; $key = [string]$f.key
        if (-not $key -or $ad -notmatch '^[A-Za-z][A-Za-z0-9-]{0,63}$') { continue }
        if ($mode -eq 'ad' -or ($mode -eq 'mixed' -and -not [string]$f.graph)) { $out += [pscustomobject]@{ key = $key; ad = $ad } }
    }
    return , $out
}

function Get-GroupsUsed($Config) {
    $g = @()
    foreach ($s in @($Config.signatures)) { foreach ($n in @($s.audience.groups)) { $l = ([string]$n).Trim().ToLowerInvariant(); if ($l -and $g -notcontains $l) { $g += $l } } }
    return , $g
}

# Every active user with a mail address: @{ mail; attrs = @{ <ad attribute> = value }; groups = @(<names, lower case>) }.
# OSIGN_TEST_AD_JSON (tests only) stands in for the directory with a JSON file of such records.
function Get-AdUserRecords([string[]]$Attributes, [string[]]$GroupNames) {
    if ($env:OSIGN_TEST_AD_JSON) {
        $out = @()
        # Assigned first: Windows PowerShell 5.1's ConvertFrom-Json hands a JSON list on as one
        # object, so @(... | ConvertFrom-Json) would be a list holding the whole list once.
        $list = (Read-Text $env:OSIGN_TEST_AD_JSON) | ConvertFrom-Json
        foreach ($r in @($list)) {
            $attrs = @{}; foreach ($p in $r.attrs.PSObject.Properties) { $attrs[$p.Name.ToLowerInvariant()] = [string]$p.Value }
            $out += @{ mail = [string]$r.mail; attrs = $attrs; groups = @(@($r.groups) | ForEach-Object { ([string]$_).ToLowerInvariant() }) }
        }
        return , $out
    }
    if (-not $script:OnWindows) { throw 'Reading users needs Windows on a machine in the domain.' }
    Add-Type -AssemblyName System.DirectoryServices -ErrorAction SilentlyContinue
    # Members of each named group, nested ones included: one search per group, not per user.
    $memberOf = @{}
    foreach ($gname in $GroupNames) {
        $v = ConvertTo-LdapFilterValue $gname
        $gs = New-Object System.DirectoryServices.DirectorySearcher
        $gs.Filter = "(&(objectClass=group)(|(sAMAccountName=$v)(name=$v)))"
        [void]$gs.PropertiesToLoad.Add('distinguishedName')
        foreach ($g in @($gs.FindAll())) {
            $dn = ConvertTo-LdapFilterValue ([string]$g.Properties['distinguishedname'][0])
            $ms = New-Object System.DirectoryServices.DirectorySearcher
            $ms.Filter = "(&(objectCategory=person)(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:=$dn))"
            $ms.PageSize = 500
            [void]$ms.PropertiesToLoad.Add('distinguishedName')
            foreach ($u in @($ms.FindAll())) {
                $udn = [string]$u.Properties['distinguishedname'][0]
                if (-not $memberOf.ContainsKey($udn)) { $memberOf[$udn] = @() }
                if ($memberOf[$udn] -notcontains $gname) { $memberOf[$udn] += $gname }
            }
        }
    }
    $us = New-Object System.DirectoryServices.DirectorySearcher
    # Users with a mail address whose account is not disabled.
    $us.Filter = '(&(objectCategory=person)(objectClass=user)(mail=*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
    $us.PageSize = 500
    foreach ($a in (@('mail', 'distinguishedName') + $Attributes)) { [void]$us.PropertiesToLoad.Add($a) }
    $out = @()
    foreach ($u in @($us.FindAll())) {
        $attrs = @{}
        foreach ($a in $Attributes) { $l = $a.ToLowerInvariant(); $attrs[$l] = $(if ($u.Properties[$l].Count) { [string]$u.Properties[$l][0] } else { '' }) }
        $dn = [string]$u.Properties['distinguishedname'][0]
        $out += @{ mail = [string]$u.Properties['mail'][0]; attrs = $attrs; groups = $(if ($memberOf.ContainsKey($dn)) { @($memberOf[$dn]) } else { @() }) }
    }
    return , $out
}

# The file name, as signature-core.js's userFileName() makes it.
function Get-UserFileName([string]$Mail) {
    $m = $Mail.Trim().ToLowerInvariant()
    if ($m -notmatch '^[^@\s]+@[^@\s]+$') { return $null }
    return (($m -replace '[^a-z0-9@._-]', '_') + '.json')
}

# What an export would do, without doing it: for every user the file it would write, and whether
# that is new, changed or the same as what is in the folder now; and the files that would go.
function New-UserSnapshot([string]$Root, $Fields, [string[]]$GroupNames) {
    $dir = Join-Path $Root 'users'
    $records = Get-AdUserRecords @($Fields | ForEach-Object { $_.ad } | Select-Object -Unique) $GroupNames
    $seen = @{}; $entries = @()
    foreach ($r in $records) {
        $name = Get-UserFileName $r.mail
        if (-not $name -or $seen.ContainsKey($name)) { continue }
        $seen[$name] = 1
        $values = [ordered]@{}
        foreach ($f in ($Fields | Sort-Object key)) { $values[$f.key] = [string]$r.attrs[$f.ad.ToLowerInvariant()] }
        $groups = @($r.groups | Sort-Object)
        $json = To-Json ([ordered]@{ values = $values; groups = $groups })
        $old = Read-Text (Join-Path $dir $name)
        $status = $(if ($null -eq $old) { 'new' } elseif ($old -eq $json) { 'unchanged' } else { 'changed' })
        $oldValues = $null
        if ($status -eq 'changed') { try { $oldValues = ($old | ConvertFrom-Json) } catch { } }
        $entries += [ordered]@{ file = $name; mail = $r.mail; status = $status; values = $values; groups = $groups; old = $oldValues; json = $json }
    }
    $removed = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
        if (-not $seen.ContainsKey($f.Name)) { $removed += $f.Name }
    }
    return @{ entries = $entries; removed = $removed }
}

# Writes what a snapshot holds: the new and changed files, and removes the ones that went. A file
# that has changed in the folder since the snapshot was taken is written over all the same: the
# snapshot is what was checked.
function Write-UserSnapshot([string]$Root, $Snapshot) {
    $dir = Join-Path $Root 'users'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $written = 0; $same = 0; $removed = 0
    foreach ($e in @($Snapshot.entries)) {
        $path = Join-Path $dir ([string]$e.file)
        if ((Read-Text $path) -eq [string]$e.json) { $same++ } else { Write-Text $path ([string]$e.json); $written++ }
    }
    foreach ($n in @($Snapshot.removed)) {
        $path = Join-Path $dir ([string]$n)
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force; $removed++ }
    }
    return @{ users = @($Snapshot.entries).Count; written = $written; unchanged = $same; removed = $removed }
}

# The scheduled export (-ExportUsers): read and write in one go, unattended.
function Export-UsersToFolder([string]$Root, $Fields, [string[]]$GroupNames) {
    return (Write-UserSnapshot $Root (New-UserSnapshot $Root $Fields $GroupNames))
}

function Read-ExportRequest($data) {
    $fields = @(); foreach ($f in @($data.fields)) { if ([string]$f.key -and ([string]$f.ad) -match '^[A-Za-z][A-Za-z0-9-]{0,63}$') { $fields += [pscustomobject]@{ key = [string]$f.key; ad = [string]$f.ad } } }
    $groups = @(@($data.groups) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    # What the snapshot was taken for: another set of fields or groups makes it another export.
    $sig = (@($fields | Sort-Object key | ForEach-Object { "$($_.key)=$($_.ad)" }) -join ';') + '|' + ($groups -join ';')
    return @{ fields = $fields; groups = $groups; sig = $sig }
}

$script:SnapshotFile = Join-Path $script:DataDir 'users-snapshot.json'

# For the UI: the users as read, without the file contents.
function Get-SnapshotSummary($snap) {
    $users = @(@($snap.entries) | ForEach-Object { [ordered]@{ mail = $_.mail; status = $_.status; values = $_.values; groups = $_.groups; old = $(if ($_.old) { $_.old.values } else { $null }) } })
    $count = { param($st) @(@($snap.entries) | Where-Object { $_.status -eq $st }).Count }
    return @{
        readAt = $snap.readAt; sig = $snap.sig; written = [bool]$snap.written; users = $users
        removed = @(@($snap.removed) | ForEach-Object { ([string]$_) -replace '\.json$', '' })
        counts = [ordered]@{ new = (& $count 'new'); changed = (& $count 'changed'); unchanged = (& $count 'unchanged'); removed = @($snap.removed).Count }
    }
}

# Read the users from AD and keep what was read, for checking; nothing is written to the folder.
function Api-AdRead($data) {
    $s = Get-Settings
    $err = Test-Settings $s -ForPublish; if ($err) { return (Fail $err) }
    $req = Read-ExportRequest $data
    try { $snap = New-UserSnapshot $s.folder $req.fields $req.groups } catch { return (Fail "Active Directory: $($_.Exception.Message)") }
    $snap.readAt = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds(); $snap.sig = $req.sig; $snap.written = $false
    Write-Text $script:SnapshotFile (To-Json $snap)
    return (Ok (Get-SnapshotSummary $snap))
}

# The last read, if there is one: { ok, snapshot: null | summary, matches }. matches: it was
# read for these fields and groups (the ones the UI sends), so Publish may write it.
function Api-AdSnapshot($data) {
    $raw = Read-Text $script:SnapshotFile
    if (-not $raw) { return (Ok @{ snapshot = $null; matches = $false }) }
    try {
        $snap = $raw | ConvertFrom-Json
        return (Ok @{ snapshot = (Get-SnapshotSummary $snap); matches = ([string]$snap.sig -eq (Read-ExportRequest $data).sig) })
    } catch { return (Ok @{ snapshot = $null; matches = $false }) }
}

# Write the users as they were read and checked. Refused when nothing was read, when it was read
# for other fields or groups, or when it was written already (then AD has to be read again: an old
# read must not overwrite what a later export wrote).
function Api-AdExport($data) {
    $s = Get-Settings
    $err = Test-Settings $s -ForPublish; if ($err) { return (Fail $err) }
    $req = Read-ExportRequest $data
    $raw = Read-Text $script:SnapshotFile
    if (-not $raw) { return (Fail 'Read the users from AD and check them first (Users).') }
    $snap = $raw | ConvertFrom-Json
    if ([string]$snap.sig -ne $req.sig) { return (Fail 'The fields or groups changed since the users were read: read them again (Users).') }
    if ($snap.written) { return (Fail 'These users were written already: read them again to write newer values (Users).') }
    try { $r = Write-UserSnapshot $s.folder $snap } catch { return (Fail "Writing the users: $($_.Exception.Message)") }
    $snap.written = $true
    Write-Text $script:SnapshotFile (To-Json $snap)
    return (Ok $r)
}

# ---------------------------------------------------------------------------
# HTTP server: 127.0.0.1 only, a token per start
# ---------------------------------------------------------------------------
$script:Token = -join ((1..32) | ForEach-Object { '{0:x2}' -f (Get-Random -Minimum 0 -Maximum 256) })
try {
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create(); $tb = New-Object byte[] 32; $rng.GetBytes($tb); $rng.Dispose()
    $script:Token = ($tb | ForEach-Object { $_.ToString('x2') }) -join ''
} catch { }

function Read-Request($client) {
    $ns = $client.GetStream(); $ns.ReadTimeout = 8000
    $ms = New-Object System.IO.MemoryStream; $buf = New-Object byte[] 65536; $headerEnd = -1; $scan = 0
    try {
        while ($true) {
            $read = $ns.Read($buf, 0, $buf.Length); if ($read -le 0) { break }
            $ms.Write($buf, 0, $read); $arr = $ms.ToArray()
            for ($i = [Math]::Max(0, $scan - 3); $i -le $arr.Length - 4; $i++) { if ($arr[$i] -eq 13 -and $arr[$i + 1] -eq 10 -and $arr[$i + 2] -eq 13 -and $arr[$i + 3] -eq 10) { $headerEnd = $i; break } }
            $scan = $arr.Length
            if ($headerEnd -lt 0) { if ($arr.Length -gt 65536) { break } else { continue } }
            $htext = [Text.Encoding]::ASCII.GetString($arr, 0, $headerEnd); $cl = 0
            if ($htext -match '(?im)^Content-Length:\s*(\d+)') { $cl = [int64]$Matches[1] }
            if ($cl -gt 8MB) { return $null }
            $bodyStart = $headerEnd + 4; $have = $arr.Length - $bodyStart
            while ($have -lt $cl) { $read = $ns.Read($buf, 0, $buf.Length); if ($read -le 0) { break }; $ms.Write($buf, 0, $read); $have += $read }
            $arr = $ms.ToArray(); $body = ''
            if ($cl -gt 0) { $body = $script:Utf8.GetString($arr, $bodyStart, [Math]::Min($cl, $arr.Length - $bodyStart)) }
            $parts = (($htext -split "`r`n")[0]) -split ' '
            $h = { param($n) if ($htext -match "(?im)^${n}:[ \t]*([^\r\n]*)") { $Matches[1].Trim() } else { $null } }
            return @{ method = $parts[0]; path = [string]$parts[1]; body = $body; host = [string](& $h 'Host'); origin = (& $h 'Origin'); ctype = [string](& $h 'Content-Type'); auth = [string](& $h 'Authorization') }
        }
    } catch { }
    return $null
}

function Send-Http($client, [string]$Status, [string]$Ctype, [byte[]]$Body, [string]$Cache = 'no-store') {
    # Never inside another site's frame, never read as another type; the page loads nothing but its
    # own files, plus images from https (an image a signature links to, shown in the preview).
    $sec = "X-Frame-Options: DENY`r`nX-Content-Type-Options: nosniff`r`nReferrer-Policy: no-referrer`r`n" +
        "Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; connect-src 'self'; frame-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'self'`r`n"
    if (-not $Body) { $Body = [byte[]]@() }
    $head = "HTTP/1.1 $Status`r`nContent-Type: $Ctype`r`nContent-Length: $($Body.Length)`r`nCache-Control: $Cache`r`n$sec" + "Connection: close`r`n`r`n"
    $hb = [Text.Encoding]::ASCII.GetBytes($head); $ns = $client.GetStream(); $ns.Write($hb, 0, $hb.Length); if ($Body.Length) { $ns.Write($Body, 0, $Body.Length) }; $ns.Flush()
}
function Send-Json($client, [string]$Json) { Send-Http $client '200 OK' 'application/json; charset=utf-8' ($script:Utf8.GetBytes($Json)) }
function Send-Status($client, [string]$Status) { Send-Http $client $Status 'text/plain; charset=utf-8' ($script:Utf8.GetBytes($Status)) }

# Compared in constant time.
function Test-ApiToken([string]$Auth) {
    if ($Auth -notmatch '^Bearer ([0-9a-f]{64})$') { return $false }
    $t = $Matches[1]; $d = 0
    for ($i = 0; $i -lt $t.Length; $i++) { $d = $d -bor ([int][char]$t[$i] -bxor [int][char]$script:Token[$i]) }
    return ($d -eq 0)
}

$script:StaticTypes = @{ '.html' = 'text/html; charset=utf-8'; '.js' = 'text/javascript; charset=utf-8'; '.css' = 'text/css; charset=utf-8'; '.svg' = 'image/svg+xml'; '.png' = 'image/png'; '.ico' = 'image/x-icon' }

function Send-File($client, [string]$Root, [string]$Rel, [hashtable]$Types) {
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $full = [IO.Path]::GetFullPath((Join-Path $Root ($Rel -replace '/', [IO.Path]::DirectorySeparatorChar)))
    $ext = [IO.Path]::GetExtension($full).ToLowerInvariant()
    if (-not $full.StartsWith($rootFull) -or -not $Types.ContainsKey($ext) -or -not (Test-Path -LiteralPath $full -PathType Leaf)) { Send-Status $client '404 Not Found'; return }
    Send-Http $client '200 OK' $Types[$ext] ([IO.File]::ReadAllBytes($full))
}

$script:LastPing = Get-Date
$script:Quit = $false

function Invoke-Request($client) {
    $req = Read-Request $client
    if (-not $req) { Send-Status $client '400 Bad Request'; return }
    # Only a request addressed to this machine by its own name: another site that makes its name
    # resolve to 127.0.0.1 (DNS rebinding) still sends its own name as the Host.
    if ($req.host -notmatch '^(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$') { Send-Status $client '403 Forbidden'; return }
    $path = ($req.path -split '\?')[0]
    try { $path = [Uri]::UnescapeDataString($path) } catch { Send-Status $client '400 Bad Request'; return }

    if ($path -like '/api/*') {
        $okOrigin = ($null -eq $req.origin) -or ($req.origin -match '^http://(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$')
        if (-not $okOrigin -or $req.method -ne 'POST' -or $req.ctype -notmatch '^application/json') { Send-Status $client '403 Forbidden'; return }
        if (-not (Test-ApiToken $req.auth)) { Send-Http $client '401 Unauthorized' 'application/json; charset=utf-8' ($script:Utf8.GetBytes((Fail 'bad token'))); return }
        $data = $null
        try { if ($req.body) { $data = $req.body | ConvertFrom-Json } } catch { Send-Json $client (Fail 'The request is not valid JSON.'); return }
        if (-not $data) { $data = [pscustomobject]@{} }
        $script:LastPing = Get-Date
        $out = switch ($path) {
            '/api/ping'           { Ok @{} }
            '/api/info'           { Api-Info }
            '/api/draft/get'      { Api-DraftGet }
            '/api/draft/save'     { Api-DraftSave $data }
            '/api/assets/upload'  { Api-AssetUpload $data }
            '/api/assets/delete'  { Api-AssetDelete $data }
            '/api/settings/get'   { Api-SettingsGet }
            '/api/settings/save'  { Api-SettingsSave $data }
            '/api/publish'        { Api-Publish $data }
            '/api/addin/build'    { Api-AddinBuild }
            '/api/addin/check'    { Api-AddinCheck }
            '/api/sp/resolve'     { Api-SpResolve $data }
            '/api/pick-folder'    { Api-PickFolder $data }
            '/api/load-published' { Api-LoadPublished }
            '/api/ad/lookup'      { Api-AdLookup $data }
            '/api/ad/groups'      { Api-AdGroups $data }
            '/api/ad/read'        { Api-AdRead $data }
            '/api/ad/snapshot'    { Api-AdSnapshot $data }
            '/api/ad/export'      { Api-AdExport $data }
            '/api/quit'           { $script:Quit = $true; Ok @{} }
            default               { $null }
        }
        if ($null -eq $out) { Send-Http $client '404 Not Found' 'application/json; charset=utf-8' ($script:Utf8.GetBytes((Fail 'unknown'))); return }
        Send-Json $client $out; return
    }
    if ($req.method -ne 'GET') { Send-Status $client '405 Method Not Allowed'; return }
    if ($path -eq '/' -or $path -eq '/index.html') { Send-File $client $script:UiDir 'index.html' $script:StaticTypes; return }
    if ($path -eq '/shared/signature-core.js') {
        if ($script:CoreFile) { Send-Http $client '200 OK' 'text/javascript; charset=utf-8' ([IO.File]::ReadAllBytes($script:CoreFile)) } else { Send-Status $client '404 Not Found' }
        return
    }
    # The draft's images, for the editor and its preview. No token: an <img> cannot send one, and
    # these are the logos that will be published for anyone to load anyway.
    if ($path -like '/assets/*') { Send-File $client $script:AssetDir $path.Substring(8) $script:AssetTypes; return }
    # The photos per person, for the preview: from the photos folder (the synced one, or a share).
    if ($path -like '/photos/*') {
        $pd = Get-PhotosDir
        if ($pd) { Send-File $client $pd $path.Substring(8) $script:AssetTypes } else { Send-Status $client '404 Not Found' }
        return
    }
    if ($path -like '/ui/*') { Send-File $client $script:UiDir $path.Substring(4) $script:StaticTypes; return }
    Send-Status $client '404 Not Found'
}

# The window's own profile, set before every start: no offer to translate the editor (it has its
# own languages), and no password manager or form filling. Chrome and Edge read this file at start.
function Set-BrowserPrefs([string]$ProfileDir) {
    try {
        $dir = Join-Path $ProfileDir 'Default'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $path = Join-Path $dir 'Preferences'
        $prefs = $null
        if (Test-Path -LiteralPath $path) { try { $prefs = (Read-Text $path) | ConvertFrom-Json } catch { $prefs = $null } }
        if (-not $prefs) { $prefs = [pscustomobject]@{} }
        $set = {
            param($obj, [string]$dotted, $value)
            $parts = $dotted -split '\.'; $cur = $obj
            for ($i = 0; $i -lt $parts.Length - 1; $i++) {
                $k = $parts[$i]
                if (-not $cur.PSObject.Properties[$k] -or $cur.$k -isnot [pscustomobject]) { $cur | Add-Member -NotePropertyName $k -NotePropertyValue ([pscustomobject]@{}) -Force }
                $cur = $cur.$k
            }
            $cur | Add-Member -NotePropertyName $parts[-1] -NotePropertyValue $value -Force
        }
        & $set $prefs 'translate.enabled' $false
        & $set $prefs 'credentials_enable_service' $false
        & $set $prefs 'profile.password_manager_enabled' $false
        & $set $prefs 'autofill.profile_enabled' $false
        & $set $prefs 'autofill.credit_card_enabled' $false
        Write-Text $path (ConvertTo-Json -InputObject $prefs -Depth 30 -Compress)
    } catch { }
}

function Start-AppWindow([string]$Url) {
    if (-not $script:OnWindows) { return $false }
    $cands = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $exe = $cands | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    if (-not $exe) { return $false }
    # A profile of its own: the process that comes back is the window's (so closing it stops the
    # server), and no extension of the everyday browser - a password manager above all - reaches it.
    $profileDir = Join-Path $env:LOCALAPPDATA 'OSign\browser'
    if (-not (Test-Path -LiteralPath $profileDir)) { New-Item -ItemType Directory -Path $profileDir -Force | Out-Null }
    Set-BrowserPrefs $profileDir
    $script:BrowserProcess = Start-Process $exe -ArgumentList @("--app=$Url", "--user-data-dir=`"$profileDir`"", '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-features=Translate,TranslateUI,msEdgeTranslate', '--window-size=1400,900') -PassThru
    return $true
}

# The functions above are what the tests load; the server starts only when the script is run.
if ($MyInvocation.InvocationName -eq '.') { return }

# -ExportUsers: the export alone, for the Task Scheduler. Exit code 0 when it worked.
if ($ExportUsers) {
    $s = Get-Settings
    $err = Test-Settings $s -ForPublish
    if ($err) { Write-Host $err -ForegroundColor Red; exit 1 }
    $cfgText = Read-Text (Join-Path $s.folder 'signatures.json')
    if (-not $cfgText) { Write-Host "No signatures.json in $($s.folder): publish from the editor first." -ForegroundColor Red; exit 1 }
    $cfg = $cfgText | ConvertFrom-Json
    if ([string]$cfg.options.valueSource -notin @('ad', 'mixed')) { Write-Host 'The signatures take their values from Entra ID: nothing to export.'; exit 0 }
    try {
        $r = Export-UsersToFolder $s.folder (Get-AdExportFields $cfg) (Get-GroupsUsed $cfg)
        # A read the editor still holds is older than what was just written: it must not go out.
        $snapRaw = Read-Text $script:SnapshotFile
        if ($snapRaw) { try { $sn = $snapRaw | ConvertFrom-Json; $sn.written = $true; Write-Text $script:SnapshotFile (To-Json $sn) } catch { } }
        Write-Host ("Exported {0} users to {1}: {2} written, {3} unchanged, {4} removed." -f $r.users, (Join-Path $s.folder 'users'), $r.written, $r.unchanged, $r.removed)
        exit 0
    } catch { Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red; exit 2 }
}

# A stable port, so the window's own settings (kept per address) stay across starts.
$listener = $null
$ports = $(if ($Port) { @($Port) } else { 17690..17697 })
foreach ($try in $ports) {
    try { $l = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $try); $l.Start(); $listener = $l; $Port = $try; break } catch { $listener = $null }
}
if (-not $listener) {
    if ($ports.Count -eq 1) { Write-Host "Port $Port is in use." -ForegroundColor Red; exit 1 }
    $listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, 0); $listener.Start(); $Port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
}
$url = "http://127.0.0.1:$Port/#t=$($script:Token)"
Write-Host ''
Write-Host "  OSign $script:AppVersion is running." -ForegroundColor Green
Write-Host "  Open:  $url"
Write-Host "  Data:  $script:DataDir"
Write-Host '  Close this window to stop the server.' -ForegroundColor DarkGray
Write-Host ''
if (-not $NoBrowser) { if (-not (Start-AppWindow $url)) { try { Start-Process $url | Out-Null } catch { } } }

while (-not $script:Quit) {
    $pending = $false
    try { $pending = $listener.Server.Poll(200000, [System.Net.Sockets.SelectMode]::SelectRead) } catch { Start-Sleep -Milliseconds 50 }
    if ($pending) {
        $client = $null
        try { $client = $listener.AcceptTcpClient() } catch { $client = $null }
        if ($client) {
            try { Invoke-Request $client }
            catch {
                Write-Host ("Request error: " + $_.Exception.Message) -ForegroundColor Yellow
                try { Send-Http $client '500 Internal Server Error' 'application/json; charset=utf-8' ($script:Utf8.GetBytes((Fail $_.Exception.Message))) } catch { }
            }
            finally { try { $client.Close() } catch { } }
        }
    }
    if ($script:BrowserProcess -and $script:BrowserProcess.HasExited) { break }
    if (((Get-Date) - $script:LastPing).TotalHours -gt 6) { break }
}
try { $listener.Stop() } catch { }
Write-Host 'Server stopped.'
