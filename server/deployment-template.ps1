<#
  NOBS OSign - where the Outlook add-in is deployed in Microsoft 365. It only reads: it changes
  nothing in Microsoft 365, in Entra ID or on the apps.

  What it does:
    1. signs in to Microsoft's add-in deployment (Integrated apps), with the PowerShell module
       O365CentralizedAddInDeployment (it offers to install it for you alone, if it is missing);
    2. reads the entry of NOBS OSign there (Get-OrganizationAddIn): deployed or not, on or off,
       its version, and who it is assigned to;
    3. sends that to your NOBS OSign server (with the one-time code below, valid for an hour),
       where it shows under Settings -> Server.

  Run it in Windows PowerShell (5.1) as a Global or Exchange administrator:
    powershell -ExecutionPolicy Bypass -File .\nobs-osign-deployment.ps1
#>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Filled in by the server.
$Server = __OSIGN_SERVER__
$Token = __OSIGN_TOKEN__
$AddinId = __OSIGN_ADDIN_ID__
# Only the tests send the result elsewhere.
if ($env:OSIGN_SETUP_TEST_SERVER) { $Server = $env:OSIGN_SETUP_TEST_SERVER }

$De = (Get-UICulture).TwoLetterISOLanguageName -eq 'de'
function Say($en, $de, $color = 'Gray') { Write-Host $(if ($De) { $de } else { $en }) -ForegroundColor $color }
function Send($body) {
    $body.token = $Token
    $json = $body | ConvertTo-Json -Depth 5 -Compress
    Invoke-RestMethod -Method Post -Uri "$Server/api/deploy/finish" -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json)) | Out-Null
}
# A property as text: lists joined, at most 500 characters.
function Text($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) { $v = (@($v) | ForEach-Object { [string]$_ }) -join ', ' }
    $s = [string]$v
    if ($s.Length -gt 500) { $s = $s.Substring(0, 500) }
    return $s
}

Say "NOBS OSign: where the add-in is deployed (only read, nothing is changed)" "NOBS OSign: wo das Add-in verteilt ist (nur lesen, nichts wird geändert)" 'Cyan'
try {
    if ($env:OSIGN_DEPLOY_TEST_ADDIN) {
        # Only the tests: what Get-OrganizationAddIn would answer, as JSON.
        $found = $env:OSIGN_DEPLOY_TEST_ADDIN | ConvertFrom-Json
    } else {
        if (-not (Get-Module -ListAvailable -Name O365CentralizedAddInDeployment)) {
            Say "The module O365CentralizedAddInDeployment is missing. Install it for you alone (Install-Module -Scope CurrentUser)? [y/n]" "Das Modul O365CentralizedAddInDeployment fehlt. Nur für Sie installieren (Install-Module -Scope CurrentUser)? [j/n]" 'Yellow'
            if ((Read-Host) -notmatch '^[jy]') { throw 'The module O365CentralizedAddInDeployment is needed.' }
            Install-Module -Name O365CentralizedAddInDeployment -Scope CurrentUser -Force
        }
        Import-Module O365CentralizedAddInDeployment
        Say "Signing in (a window opens)..." "Anmelden (ein Fenster öffnet sich)..."
        Connect-OrganizationAddInService | Out-Null
        $found = $null
        try { $found = Get-OrganizationAddIn -ProductId $AddinId } catch { $found = $null }
    }
    $props = @{}
    if ($found) {
        foreach ($p in @($found.PSObject.Properties) | Select-Object -First 40) { $props[$p.Name] = Text $p.Value }
        Say ("Found: " + $props['DisplayName'] + " " + $props['Version']) ("Gefunden: " + $props['DisplayName'] + " " + $props['Version']) 'Green'
    } else {
        Say "NOBS OSign is not deployed (Integrated apps)." "NOBS OSign ist nicht verteilt (Integrierte Apps)." 'Yellow'
    }
    Send @{ found = [bool]$found; props = $props }
    Say "Sent to NOBS OSign ($Server): see Settings -> Server." "An NOBS OSign übermittelt ($Server): siehe Einstellungen -> Server." 'Green'
} catch {
    $why = $_.Exception.Message
    Say "Stopped: $why" "Abgebrochen: $why" 'Red'
    try { Send @{ error = $why } } catch { }
    exit 1
}
