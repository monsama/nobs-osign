<#
  NOBS OSign - setup for Microsoft 365, made by the setup assistant for one setup.

  What it does, signed in once with your administrator account (Microsoft's sign-in in the browser):
    1. the add-in's app registration "NOBS OSign": single-page application, redirect URI
       brk-multihub://<add-in host>, Microsoft Graph delegated User.Read, Sites.Read.All (and
       GroupMember.Read.All where groups come from Entra ID), with admin consent;
    2. the server's app registration "NOBS OSign server": application permission Sites.Selected
       with admin consent, and the server's certificate on it (the server keeps the key; where
       that is refused, a client secret instead);
    3. write access for "NOBS OSign server" on the one SharePoint site of the signatures;
    4. sends the client IDs (and the secret, if one was made) to your NOBS OSign server (with the one-time code
       below, valid for an hour), so nothing has to be copied;
    5. names apps of yours by these names left over from earlier runs (whose result never reached
       the server; the newest is used again), and removes them if you say so;
    6. takes back the rights it was given for this (the consent for "Microsoft Graph Command Line
       Tools", the sign-in it uses), and deletes itself.

  Run it on a Windows PC as the SharePoint or Global administrator:
    powershell -ExecutionPolicy Bypass -File .\nobs-osign-setup.ps1
  It can run again: the apps it made before (as the server knows them, and owned by you) are used
  again; otherwise new ones are made.
#>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Filled in by the server for this setup.
$Server = __OSIGN_SERVER__
$Token = __OSIGN_TOKEN__
$Site = __OSIGN_SITE__
$AddinHost = __OSIGN_ADDIN_HOST__
$WantGroups = __OSIGN_GROUPS__
# The server may read users and their groups from Entra ID too (the preview of a real person).
$WantUsers = __OSIGN_USERS__
# The apps of an earlier run (client IDs the server has): used again if you own them. Any other
# app is never used, whatever its name: anyone may register an app called "NOBS OSign server".
$KnownAddin = __OSIGN_KNOWN_ADDIN__
$KnownServer = __OSIGN_KNOWN_SERVER__
# The server's certificate (only the public half, base64 DER; the key stays on the server).
$Cert = __OSIGN_CERT__
$CertThumbprint = __OSIGN_CERT_THUMBPRINT__

# Microsoft's endpoints; only the tests point them elsewhere.
# Only the tests send the result elsewhere (to a server that is not there).
if ($env:OSIGN_SETUP_TEST_SERVER) { $Server = $env:OSIGN_SETUP_TEST_SERVER }
$LoginUrl = if ($env:OSIGN_SETUP_LOGIN_URL) { $env:OSIGN_SETUP_LOGIN_URL } else { 'https://login.microsoftonline.com' }
$GraphUrl = if ($env:OSIGN_SETUP_GRAPH_URL) { $env:OSIGN_SETUP_GRAPH_URL } else { 'https://graph.microsoft.com/v1.0' }
# The sign-in: Microsoft's own public client for Graph scripts ("Microsoft Graph Command Line
# Tools", the one Connect-MgGraph uses), so no app of ours needs these rights.
$CliApp = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$GraphApp = '00000003-0000-0000-c000-000000000000'
$Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'Sites.FullControl.All')
$AddinName = 'NOBS OSign'
$ServerName = 'NOBS OSign server'

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
$German = (Get-Culture).TwoLetterISOLanguageName -eq 'de'
function Say([string]$En, [string]$De, [string]$Color = 'Gray') { Write-Host $(if ($German) { $De } else { $En }) -ForegroundColor $Color }
function Step([string]$En, [string]$De) { Say ('- ' + $En) ('- ' + $De) 'Cyan' }

function B64Url([byte[]]$Bytes) { [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
function Random-B64([int]$Count) { $b = New-Object byte[] $Count; [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b); B64Url $b }
function Claims([string]$Jwt) {
    $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
    while ($p.Length % 4) { $p += '=' }
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}
function Error-Text($Err) {
    if ($Err.ErrorDetails -and $Err.ErrorDetails.Message) {
        try {
            $j = $Err.ErrorDetails.Message | ConvertFrom-Json
            if ($j.error.message) { return [string]$j.error.message }
            if ($j.error_description) { return ([string]$j.error_description).Split("`n")[0].Trim() }
        } catch { }
    }
    return $Err.Exception.Message
}

# --- Sign-in: authorization code with PKCE, back to a port on this PC --------------------------
function Sign-In {
    $l4 = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0)
    $l4.Start()
    $port = $l4.LocalEndpoint.Port
    $listeners = @($l4)
    try { $l6 = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::IPv6Loopback, $port); $l6.Start(); $listeners += $l6 } catch { }
    $redirect = "http://localhost:$port"
    $verifier = Random-B64 32
    $challenge = B64Url ([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))
    $state = Random-B64 16
    $scope = (($Scopes | ForEach-Object { 'https://graph.microsoft.com/' + $_ }) + 'openid', 'profile') -join ' '
    $q = @{ client_id = $CliApp; response_type = 'code'; redirect_uri = $redirect; scope = $scope; code_challenge = $challenge; code_challenge_method = 'S256'; state = $state; prompt = 'select_account' }
    $url = "$LoginUrl/organizations/oauth2/v2.0/authorize?" + (($q.GetEnumerator() | ForEach-Object { $_.Key + '=' + [Uri]::EscapeDataString($_.Value) }) -join '&')
    Say 'A browser window opens: sign in with your administrator account, and accept the consent (do not tick "Consent on behalf of your organization").' `
        'Ein Browserfenster öffnet sich: mit Ihrem Administratorkonto anmelden und zustimmen ("Im Namen Ihrer Organisation zustimmen" nicht anhaken).' 'Yellow'
    if ($env:OSIGN_SETUP_TEST_FETCH) { $job = Start-Job -ScriptBlock { param($u) try { Invoke-WebRequest -Uri $u -UseBasicParsing -MaximumRedirection 5 | Out-Null } catch { } } -ArgumentList $url }
    else {
        try { Start-Process $url } catch { Say "Open this address in the browser: $url" "Diese Adresse im Browser öffnen: $url" 'Yellow' }
    }
    $answer = $null
    $until = (Get-Date).AddMinutes(5)
    try {
        while (-not $answer -and (Get-Date) -lt $until) {
            $l = $listeners | Where-Object { $_.Pending() } | Select-Object -First 1
            if (-not $l) { Start-Sleep -Milliseconds 200; continue }
            $c = $l.AcceptTcpClient()
            try {
                $s = $c.GetStream(); $s.ReadTimeout = 2000
                $buf = New-Object byte[] 16384; $text = ''; $read = [Diagnostics.Stopwatch]::StartNew()
                # At most 16 KB and 5 seconds per request: nothing on this PC can hold it open.
                while ($text -notmatch "`r`n`r`n" -and $text.Length -lt 16384 -and $read.ElapsedMilliseconds -lt 5000) {
                    $n = 0; try { $n = $s.Read($buf, 0, $buf.Length) } catch { break }
                    if ($n -le 0) { break }
                    $text += [Text.Encoding]::ASCII.GetString($buf, 0, $n)
                }
                $first = ($text -split "`r`n")[0]
                $p = @{}
                if ($first -match '^GET /\?([^ ]*) HTTP') {
                    foreach ($kv in $Matches[1].Split('&')) { $i = $kv.IndexOf('='); if ($i -gt 0) { $p[$kv.Substring(0, $i)] = [Uri]::UnescapeDataString($kv.Substring($i + 1).Replace('+', ' ')) } }
                }
                $ok = $p['state'] -eq $state -and ($p['code'] -or $p['error'])
                $msg = if ($German) { 'NOBS OSign: Sie können dieses Fenster schliessen und zu PowerShell zurückkehren.' } else { 'NOBS OSign: you can close this window and go back to PowerShell.' }
                $body = [Text.Encoding]::UTF8.GetBytes('<!doctype html><meta charset="utf-8"><title>NOBS OSign</title><p style="font:16px system-ui;margin:40px">' + $msg + '</p>')
                $status = if ($ok) { '200 OK' } else { '404 Not Found' }
                $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $status`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
                $s.Write($head, 0, $head.Length); $s.Write($body, 0, $body.Length); $s.Flush()
                if ($ok) { $answer = $p }
            } finally { $c.Close() }
        }
    } finally { $listeners | ForEach-Object { $_.Stop() }; if ($job) { Remove-Job $job -Force -ErrorAction SilentlyContinue } }
    if (-not $answer) { throw $(if ($German) { 'Keine Anmeldung innert 5 Minuten.' } else { 'No sign-in within 5 minutes.' }) }
    if ($answer['error']) { throw ("Microsoft: " + $answer['error_description'] + ' (' + $answer['error'] + ')') }
    try {
        $t = Invoke-RestMethod -Method Post -Uri "$LoginUrl/organizations/oauth2/v2.0/token" -ContentType 'application/x-www-form-urlencoded' `
            -Body @{ client_id = $CliApp; grant_type = 'authorization_code'; code = $answer['code']; redirect_uri = $redirect; code_verifier = $verifier; scope = $scope }
    } catch { throw ("Microsoft: " + (Error-Text $_)) }
    return $t.access_token
}

# --- Microsoft Graph ------------------------------------------------------------------------------
function Graph([string]$Method, [string]$Path, $Body = $null, [int]$Retries = 0) {
    for ($i = 0; ; $i++) {
        try {
            $p = @{ Method = $Method; Uri = ($GraphUrl + $Path); Headers = @{ Authorization = "Bearer $script:AccessToken" } }
            if ($null -ne $Body) { $p.ContentType = 'application/json; charset=utf-8'; $p.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress)) }
            return Invoke-RestMethod @p
        } catch {
            $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            # A new app or service principal takes a moment to be known everywhere.
            if ($i -lt $Retries -and ($code -eq 400 -or $code -eq 404)) { Start-Sleep -Seconds 5; continue }
            $e = New-Object Exception ("Microsoft Graph ($Method $($Path.Split('?')[0])): " + (Error-Text $_))
            $e.Data['status'] = $code
            throw $e
        }
    }
}
function Graph-OrNull([string]$Path) { try { Graph GET $Path } catch { if ($_.Exception.Data['status'] -eq 404) { return $null }; throw } }
function Filter([string]$F) { '?$filter=' + [Uri]::EscapeDataString($F) }

# An app of an earlier run, by its client ID: only if you are one of its owners.
function Known-App([string]$AppId, [string]$Name) {
    if (-not $AppId) { return $null }
    $a = Graph-OrNull "/applications(appId='$AppId')?`$select=id,appId,displayName,spa,requiredResourceAccess"
    if (-not $a) { return $null }
    $owners = @((Graph GET "/applications/$($a.id)/owners?`$select=id").value)
    if ($owners | Where-Object { $_.id -eq $script:MyId }) { return $a }
    Say "The app ""$($a.displayName)"" ($AppId) is not yours: it is not used." "Die App ""$($a.displayName)"" ($AppId) gehört nicht Ihnen: sie wird nicht verwendet." 'Yellow'
    return $null
}
# Yours by this name, the newest first: apps of an earlier run whose result never reached the
# server (it does not know them). Another's app of the same name is never used.
function Mine-By-Name([string]$Name) {
    $all = @((Graph GET ('/applications' + (Filter "displayName eq '$Name'") + '&$select=id,appId,displayName,spa,requiredResourceAccess,createdDateTime')).value)
    @($all | Where-Object { $_ -and (@((Graph GET "/applications/$($_.id)/owners?`$select=id").value) | Where-Object { $_.id -eq $script:MyId }) } |
        Sort-Object { [string]$_.createdDateTime } -Descending)
}
# The app to use: the one the server knows, else your newest of that name. The rest of yours by that
# name are left over: named at the end, and removed if you say so.
$script:Leftovers = @()
function Find-App([string]$AppId, [string]$Name) {
    $known = Known-App $AppId $Name
    $mine = @(Mine-By-Name $Name)
    $use = if ($known) { $known } else { $mine | Select-Object -First 1 }
    if (-not $known -and $use) { Step "App registration ""$Name"" of an earlier run found ($($use.appId)): used again" "App-Registrierung ""$Name"" eines früheren Laufs gefunden ($($use.appId)): wird wieder verwendet" }
    $script:Leftovers += @($mine | Where-Object { $_ -and (-not $use -or $_.appId -ne $use.appId) })
    return $use
}
function Service-Principal([string]$AppId) {
    $sp = Graph-OrNull "/servicePrincipals(appId='$AppId')?`$select=id,appId"
    if (-not $sp) { $sp = Graph POST '/servicePrincipals' @{ appId = $AppId } -Retries 6 }
    return $sp
}
# The app's Graph permissions, with these added (and what else it asks for kept).
function With-Graph-Access($App, [array]$Access) {
    $other = @($App.requiredResourceAccess | Where-Object { $_ -and $_.resourceAppId -ne $GraphApp })
    $mine = @($App.requiredResourceAccess | Where-Object { $_ -and $_.resourceAppId -eq $GraphApp } | ForEach-Object { $_.resourceAccess })
    foreach ($a in $Access) { if (-not ($mine | Where-Object { $_.id -eq $a.id -and $_.type -eq $a.type })) { $mine += [pscustomobject]$a } }
    return @($other) + @(@{ resourceAppId = $GraphApp; resourceAccess = @($mine | ForEach-Object { @{ id = $_.id; type = $_.type } }) })
}

$script:AccessToken = $null
$me = $null
$result = $null
try {
    Say '' ''
    Say 'NOBS OSign - setup for Microsoft 365' 'NOBS OSign - Einrichtung für Microsoft 365' 'White'
    Say "Server: $Server   Site: $Site   Add-in: $AddinHost" "Server: $Server   Site: $Site   Add-in: $AddinHost"
    Say '' ''
    $script:AccessToken = Sign-In
    $tenant = (Claims $script:AccessToken).tid
    $script:MyId = (Claims $script:AccessToken).oid
    $me = Graph GET '/me?$select=id,userPrincipalName'
    if ($me.id) { $script:MyId = $me.id }
    Step "Signed in as $($me.userPrincipalName) (tenant $tenant)" "Angemeldet als $($me.userPrincipalName) (Mandant $tenant)"

    # The site first: without access to it nothing else is worth doing (no apps made in vain).
    $m = [regex]::Match($Site, '^([A-Za-z0-9.-]+):/(sites|teams|personal)/([^:/]+):$')
    $sitePath = if ($m.Success) { '/sites/' + $m.Groups[1].Value + ':/' + $m.Groups[2].Value + '/' + [Uri]::EscapeDataString($m.Groups[3].Value) + ':' } else { '/sites/root' }
    $personal = $m.Success -and $m.Groups[2].Value -eq 'personal'
    try { $siteInfo = Graph GET ($sitePath + '?$select=id,webUrl') }
    catch {
        if ($_.Exception.Data['status'] -eq 403 -and $personal) {
            throw $(if ($German) { "Kein Zugriff auf das OneDrive $Site mit $($me.userPrincipalName): Auf ein persönliches OneDrive kommt nur dessen Besitzer, auch kein Administrator. Legen Sie die Signaturen besser in eine SharePoint-Website (empfohlen) und wählen Sie diese im Einrichtungsassistenten." } else { "No access to the OneDrive $Site as $($me.userPrincipalName): only its owner reaches a personal OneDrive, no administrator does. Better keep the signatures in a SharePoint site (recommended) and choose it in the setup assistant." })
        }
        if ($_.Exception.Data['status'] -eq 403) {
            throw $(if ($German) { "Kein Zugriff auf die Website $Site mit $($me.userPrincipalName): Ihr Konto muss Besitzer der Website sein (in SharePoint: Website-Einstellungen → Websiteberechtigungen), dann das Skript nochmals ausführen." } else { "No access to the site $Site as $($me.userPrincipalName): your account has to be an owner of the site (in SharePoint: Site settings → Site permissions); then run the script again." })
        }
        throw
    }
    Step "Site found: $($siteInfo.webUrl)" "Website gefunden: $($siteInfo.webUrl)"

    $graphSp = Graph GET "/servicePrincipals(appId='$GraphApp')?`$select=id,appRoles,oauth2PermissionScopes"
    $scopeId = @{}; foreach ($s in $graphSp.oauth2PermissionScopes) { $scopeId[$s.value] = $s.id }
    $roleId = ($graphSp.appRoles | Where-Object { $_.value -eq 'Sites.Selected' } | Select-Object -First 1).id

    # 1. The add-in's app.
    $want = @('User.Read', 'Sites.Read.All'); if ($WantGroups) { $want += 'GroupMember.Read.All' }
    $redirect = "brk-multihub://$AddinHost"
    $app = Find-App $KnownAddin $AddinName
    $access = @($want | ForEach-Object { @{ id = $scopeId[$_]; type = 'Scope' } })
    if ($app) {
        $uris = @($app.spa.redirectUris) + $redirect | Where-Object { $_ } | Select-Object -Unique
        Graph PATCH "/applications/$($app.id)" @{ spa = @{ redirectUris = @($uris) }; requiredResourceAccess = @(With-Graph-Access $app $access) } | Out-Null
        Step "App registration ""$AddinName"" is there: redirect URI and permissions checked" "App-Registrierung ""$AddinName"" ist da: Umleitungs-URI und Berechtigungen geprüft"
    } else {
        $app = Graph POST '/applications' @{ displayName = $AddinName; signInAudience = 'AzureADMyOrg'; spa = @{ redirectUris = @($redirect) }; requiredResourceAccess = @(@{ resourceAppId = $GraphApp; resourceAccess = $access }) }
        Step "App registration ""$AddinName"" made" "App-Registrierung ""$AddinName"" angelegt"
    }
    $appSp = Service-Principal $app.appId
    $grant = @((Graph GET ('/oauth2PermissionGrants' + (Filter "clientId eq '$($appSp.id)' and consentType eq 'AllPrincipals' and resourceId eq '$($graphSp.id)'"))).value) | Select-Object -First 1
    if ($grant) {
        $scope = (@($grant.scope -split ' ') + $want | Where-Object { $_ } | Select-Object -Unique) -join ' '
        if ($scope -ne $grant.scope) { Graph PATCH "/oauth2PermissionGrants/$($grant.id)" @{ scope = $scope } | Out-Null }
    } else {
        Graph POST '/oauth2PermissionGrants' @{ clientId = $appSp.id; consentType = 'AllPrincipals'; resourceId = $graphSp.id; scope = ($want -join ' ') } -Retries 6 | Out-Null
    }
    Step ("Admin consent for " + ($want -join ', ')) ("Administratorzustimmung für " + ($want -join ', '))

    # 2. The server's app.
    $srv = Find-App $KnownServer $ServerName
    $role = @{ id = $roleId; type = 'Role' }
    if ($srv) {
        Graph PATCH "/applications/$($srv.id)" @{ requiredResourceAccess = @(With-Graph-Access $srv @($role)) } | Out-Null
        Step "App registration ""$ServerName"" is there" "App-Registrierung ""$ServerName"" ist da"
    } else {
        $srv = Graph POST '/applications' @{ displayName = $ServerName; signInAudience = 'AzureADMyOrg'; requiredResourceAccess = @(@{ resourceAppId = $GraphApp; resourceAccess = @($role) }) }
        Step "App registration ""$ServerName"" made" "App-Registrierung ""$ServerName"" angelegt"
    }
    $srvSp = Service-Principal $srv.appId
    $roles = @('Sites.Selected'); if ($WantUsers) { $roles += 'User.Read.All', 'GroupMember.Read.All' }
    if ($WantUsers) { Graph PATCH "/applications/$($srv.id)" @{ requiredResourceAccess = @(With-Graph-Access $srv @($roles | ForEach-Object { $n = $_; @{ id = ($graphSp.appRoles | Where-Object { $_.value -eq $n } | Select-Object -First 1).id; type = 'Role' } })) } | Out-Null }
    $have = @((Graph GET "/servicePrincipals/$($srvSp.id)/appRoleAssignments").value)
    foreach ($n in $roles) {
        $rid = ($graphSp.appRoles | Where-Object { $_.value -eq $n } | Select-Object -First 1).id
        if (-not ($have | Where-Object { $_.appRoleId -eq $rid -and $_.resourceId -eq $graphSp.id })) { Graph POST "/servicePrincipals/$($srvSp.id)/appRoleAssignments" @{ principalId = $srvSp.id; resourceId = $graphSp.id; appRoleId = $rid } -Retries 6 | Out-Null }
    }
    Step ("Admin consent for " + ($roles -join ', ') + " (application)") ("Administratorzustimmung für " + ($roles -join ', ') + " (Anwendung)")
    # The server signs in with its certificate: it goes on the app in place of the ones before.
    $secret = $null; $certOn = $false
    if ($Cert) {
        try {
            Graph PATCH "/applications/$($srv.id)" @{ keyCredentials = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = $Cert; displayName = 'NOBS OSign server ' + (Get-Date -Format 'yyyy-MM-dd') }) } | Out-Null
            $certOn = $true
            Step "The server's certificate is on ""$ServerName"" (thumbprint $CertThumbprint)" "Das Zertifikat des Servers ist auf ""$ServerName"" (Fingerabdruck $CertThumbprint)"
            # Client secrets of earlier runs are not needed any more.
            $old = @((Graph GET "/applications/$($srv.id)?`$select=passwordCredentials").passwordCredentials) | Where-Object { $_ -and ([string]$_.displayName).StartsWith($ServerName) }
            foreach ($p in $old) { try { Graph POST "/applications/$($srv.id)/removePassword" @{ keyId = $p.keyId } | Out-Null } catch { } }
            if ($old) { Step ("Old client secrets removed: " + @($old).Count) ("Alte geheime Clientschlüssel entfernt: " + @($old).Count) }
        } catch {
            Say ("The certificate could not be put on the app (" + (Error-Text $_) + "): a client secret instead.") ("Das Zertifikat konnte nicht auf die App (" + (Error-Text $_) + "): stattdessen ein geheimer Clientschlüssel.") 'Yellow'
        }
    }
    if (-not $certOn) {
        foreach ($months in 24, 12, 6) {
            try { $secret = Graph POST "/applications/$($srv.id)/addPassword" @{ passwordCredential = @{ displayName = 'NOBS OSign server ' + (Get-Date -Format 'yyyy-MM-dd'); endDateTime = (Get-Date).ToUniversalTime().AddMonths($months).ToString('yyyy-MM-ddTHH:mm:ssZ') } }; break }
            catch { if ($months -eq 6) { throw } }
        }
        Step ("Client secret made, valid until " + ([datetime]$secret.endDateTime).ToString('yyyy-MM-dd')) ("Geheimer Clientschlüssel erstellt, gültig bis " + ([datetime]$secret.endDateTime).ToString('yyyy-MM-dd'))
    }

    # 3. The site (found at the start).
    $s = $siteInfo
    $perms = @((Graph GET "/sites/$($s.id)/permissions").value)
    $has = $perms | Where-Object {
        $p = $_
        (@($p.roles) | Where-Object { @('write', 'owner', 'fullcontrol') -contains ([string]$_).ToLower() }) -and
        (@($p.grantedToIdentitiesV2) + @($p.grantedToIdentities) | Where-Object { $_ -and $_.application -and $_.application.id -eq $srv.appId })
    }
    if (-not $has) { Graph POST "/sites/$($s.id)/permissions" @{ roles = @('write'); grantedToIdentities = @(@{ application = @{ id = $srv.appId; displayName = $ServerName } }) } -Retries 6 | Out-Null }
    Step "Write access for ""$ServerName"" on $($s.webUrl)" "Schreibzugriff für ""$ServerName"" auf $($s.webUrl)"

    $srvResult = if ($certOn) { @{ clientId = $srv.appId; certThumbprint = $CertThumbprint } } else { @{ clientId = $srv.appId; secret = $secret.secretText; secretExpires = $(if ($secret.endDateTime -is [datetime]) { $secret.endDateTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } else { [string]$secret.endDateTime }) } }
    $result = @{ token = $Token; tenantId = $tenant; siteUrl = $s.webUrl; addin = @{ clientId = $app.appId }; server = $srvResult }
} catch {
    $failure = $_.Exception.Message
    Say '' ''
    Say "Stopped: $failure" "Abgebrochen: $failure" 'Red'
    try { Invoke-RestMethod -Method Post -Uri "$Server/api/setup/finish" -ContentType 'application/json' -Body (@{ token = $Token; error = $failure } | ConvertTo-Json -Compress) | Out-Null } catch { }
}

# 4. To the server.
if ($result) {
    try {
        Invoke-RestMethod -Method Post -Uri "$Server/api/setup/finish" -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes(($result | ConvertTo-Json -Depth 5 -Compress))) | Out-Null
        Step "Sent to NOBS OSign ($Server): the setup assistant goes on by itself" "An NOBS OSign übermittelt ($Server): der Einrichtungsassistent geht von selbst weiter"
    } catch {
        $why = Error-Text $_
        Say ("The server could not be reached: " + $why) ("Der Server war nicht erreichbar: " + $why) 'Red'
        if ($why -match 'SSL|TLS|trust|Vertrauens') {
            Say "  This PC does not trust the certificate of $Server (it comes from your own CA, or the proxy's). The file below does not need that." `
                "  Dieser PC vertraut dem Zertifikat von $Server nicht (es stammt von Ihrer eigenen CA oder vom Proxy). Die Datei unten braucht das nicht." 'Yellow'
        }
        # The result as a file, for the browser to hand over: it reaches the server where this PC cannot.
        $file = Join-Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) 'nobs-osign-setup-result.json'
        [IO.File]::WriteAllText($file, ($result | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
        Say "Saved the result: $file" "Ergebnis gespeichert: $file" 'Yellow'
        Say '  In NOBS OSign: Settings → Server connections → SharePoint (or the setup assistant), "Upload the result" and choose this file. Valid for an hour; delete it afterwards.' `
            '  In NOBS OSign: Einstellungen → Server-Verbindungen → SharePoint (oder der Einrichtungsassistent), «Ergebnis hochladen» und diese Datei wählen. Eine Stunde gültig; danach löschen.' 'Yellow'
    }
}

# 5. Apps left over from earlier runs: named, and removed if you say so (Entra keeps them 30 days
#    under Deleted applications).
if ($result -and $script:Leftovers.Count) {
    Say '' ''
    Say 'Left over from earlier runs, not used:' 'Übrig von früheren Läufen, nicht verwendet:' 'Yellow'
    foreach ($a in $script:Leftovers) { Say "  $($a.displayName)  $($a.appId)  $($a.createdDateTime)" "  $($a.displayName)  $($a.appId)  $($a.createdDateTime)" 'Yellow' }
    $answer = if ($env:OSIGN_SETUP_TEST_FETCH) { [string]$env:OSIGN_SETUP_TEST_REMOVE } else { Read-Host $(if ($German) { 'Jetzt entfernen? (j/N)' } else { 'Remove them now? (y/N)' }) }
    if ($answer -match '^\s*(y|yes|j|ja)\s*$') {
        foreach ($a in $script:Leftovers) {
            try { Graph DELETE "/applications/$($a.id)" | Out-Null; Step "Removed: $($a.displayName) ($($a.appId))" "Entfernt: $($a.displayName) ($($a.appId))" }
            catch { Say ("Not removed: $($a.appId): " + $_.Exception.Message) ("Nicht entfernt: $($a.appId): " + $_.Exception.Message) 'Yellow' }
        }
    } else {
        Say '  Kept. Remove them in Entra ID → App registrations when they are not needed.' '  Behalten. In Entra ID → App-Registrierungen entfernen, wenn sie nicht gebraucht werden.' 'Yellow'
    }
}

# 6. The rights used for this go again: the consent given at the sign-in, for your account.
if ($script:AccessToken -and $script:MyId) {
    try {
        $cli = Graph-OrNull "/servicePrincipals(appId='$CliApp')?`$select=id"
        if ($cli) {
            $graphSpId = (Graph GET "/servicePrincipals(appId='$GraphApp')?`$select=id").id
            foreach ($g in @((Graph GET ('/oauth2PermissionGrants' + (Filter "clientId eq '$($cli.id)' and resourceId eq '$graphSpId'"))).value)) {
                $left = @($g.scope -split ' ' | Where-Object { $_ -and $Scopes -notcontains $_ })
                if ($left.Count -eq @($g.scope -split ' ' | Where-Object { $_ }).Count) { continue }
                if ($g.consentType -ne 'Principal' -or $g.principalId -ne $script:MyId) {
                    Say 'Note: these rights were given to "Microsoft Graph Command Line Tools" for the whole organization; they stay (take them back in Entra → Enterprise applications if they are not needed).' `
                        'Hinweis: diese Rechte wurden "Microsoft Graph Command Line Tools" für die ganze Organisation erteilt; sie bleiben (in Entra → Unternehmensanwendungen entfernen, wenn sie nicht gebraucht werden).' 'Yellow'
                    continue
                }
                # Nothing but the sign-in's own (openid, profile) left: the grant goes as a whole.
                if (@($left | Where-Object { @('openid', 'profile', 'offline_access', 'email') -notcontains $_ }).Count) {
                    # Where Graph refuses to shorten the grant (400), it goes and comes back with the rest.
                    try { Graph PATCH "/oauth2PermissionGrants/$($g.id)" @{ scope = ($left -join ' ') } | Out-Null }
                    catch {
                        Graph DELETE "/oauth2PermissionGrants/$($g.id)" | Out-Null
                        Graph POST '/oauth2PermissionGrants' @{ clientId = $g.clientId; consentType = 'Principal'; principalId = $g.principalId; resourceId = $g.resourceId; scope = ($left -join ' ') } | Out-Null
                    }
                } else { Graph DELETE "/oauth2PermissionGrants/$($g.id)" | Out-Null }
            }
            Step 'The rights used for this are taken back' 'Die dafür verwendeten Rechte sind wieder entfernt'
        }
    } catch {
        Say ("The rights used for this could not be taken back: " + $_.Exception.Message) ("Die verwendeten Rechte konnten nicht entfernt werden: " + $_.Exception.Message) 'Yellow'
        Say ('Take them back by hand: Entra admin center → Enterprise applications → "Microsoft Graph Command Line Tools" → Permissions → User consent: ' + ($Scopes -join ', ') + '.') `
            ('Von Hand entfernen: Entra Admin Center → Unternehmensanwendungen → "Microsoft Graph Command Line Tools" → Berechtigungen → Benutzereinwilligung: ' + ($Scopes -join ', ') + '.') 'Yellow'
    }
}
$script:AccessToken = $null

# It was made for this one setup: it goes.
if ($PSCommandPath -and -not $env:OSIGN_SETUP_KEEP) { Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue }
Say '' ''
if ($result) { Say 'Done.' 'Fertig.' 'Green' }
if (-not $env:OSIGN_SETUP_TEST_FETCH) { Read-Host $(if ($German) { 'Enter zum Schliessen' } else { 'Enter to close' }) | Out-Null }
if (-not $result) { exit 1 }
