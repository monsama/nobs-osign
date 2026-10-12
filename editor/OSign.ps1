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
#>
param(
    [switch]$NoBrowser,
    [int]$Port = 0
)

$ErrorActionPreference = 'Stop'
$script:AppVersion = '0.1.0'
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
function Get-DefaultSettings { return [ordered]@{ folder = ''; photosFolder = '' } }

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
    foreach ($k in @('folder', 'photosFolder')) {
        $v = [string]$s[$k]
        if ($v -and ($v.IndexOfAny([IO.Path]::GetInvalidPathChars()) -ge 0 -or -not [IO.Path]::IsPathRooted($v))) { return "The $k is a full path, like C:\Users\me\Contoso\Documents\OSign." }
    }
    if ($ForPublish) {
        if (-not $s.folder) { return 'Pick the folder to publish to in Settings: the OSign folder of SharePoint, synced to this PC.' }
        if (-not (Test-Path -LiteralPath $s.folder -PathType Container)) { return "The folder $($s.folder) is not there. Is the SharePoint library synced to this PC?" }
    }
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
function Api-Info {
    $ad = $false
    if ($script:OnWindows) {
        try { Add-Type -AssemblyName System.DirectoryServices -ErrorAction SilentlyContinue; $ad = [bool]([DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()) } catch { $ad = $false }
    }
    return (Ok @{ backend = 'powershell'; version = $script:AppVersion; dataDir = $script:DataDir; features = @{ adLookup = $ad } })
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

function Api-SettingsSave($data) {
    $s = Get-DefaultSettings
    foreach ($k in @($s.Keys)) { if ($null -ne $data.settings.$k) { $s[$k] = ([string]$data.settings.$k).Trim().TrimEnd('\', '/') } }
    $err = Test-Settings $s; if ($err) { return (Fail $err) }
    Write-Text $script:SettingsFile (To-Json $s)
    return (Api-SettingsGet)
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
# HTTP server (the same small server as NOBS SQL Editor's: 127.0.0.1 only, a token per start)
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
            '/api/load-published' { Api-LoadPublished }
            '/api/ad/lookup'      { Api-AdLookup $data }
            '/api/ad/groups'      { Api-AdGroups $data }
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
    $script:BrowserProcess = Start-Process $exe -ArgumentList @("--app=$Url", "--user-data-dir=`"$profileDir`"", '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-features=Translate', '--window-size=1400,900') -PassThru
    return $true
}

# The functions above are what the tests load; the server starts only when the script is run.
if ($MyInvocation.InvocationName -eq '.') { return }

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
