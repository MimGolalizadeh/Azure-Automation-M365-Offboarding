<#
.SYNOPSIS
    Step 02 | Convert Mailbox to Shared
.DESCRIPTION
    Minimal offboarding step to convert a user's Exchange Online mailbox to a
    shared mailbox.

    This step must run before license removal.
    It does not grant mailbox delegation, remove licenses, or configure any
    extra settings.
.PARAMETER UserPrincipalName
    UPN of the mailbox to convert.
.PARAMETER ExchangeOrganization
    Exchange Online organization value, for example: contoso.onmicrosoft.com
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Safe simulation mode for Azure Automation testing.
.NOTES
    Requires:
      - ExchangeOnlineManagement
      - Az.Accounts
      - Exchange.ManageAsApp / managed identity access in Exchange Online
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$ExchangeOrganization,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-02_$(Get-Date -Format 'yyyy-MM-dd').log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ExchangeOrganization) -and (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
    foreach ($name in @('ExchangeOrganization', 'EXCHANGEORGANIZATION', 'ExchangeOnlineOrganization')) {
        try {
            $val = Get-AutomationVariable -Name $name -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $ExchangeOrganization = "$val"
                break
            }
        }
        catch {}
    }
}

$ExchangeOrganization = "$ExchangeOrganization".Trim().Trim('"').Trim("'")
if ([string]::IsNullOrWhiteSpace($ExchangeOrganization)) {
    throw "ExchangeOrganization is required. Pass -ExchangeOrganization or create an Automation Variable named 'ExchangeOrganization'."
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [STEP-02] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-02] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
}

function Get-ExchangeAccessToken {
    if (-not (Get-Command -Name Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        throw 'Az.Accounts is required. Install/import Az.Accounts and sign in with Connect-AzAccount, or run inside Azure Automation with managed identity.'
    }

    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $azContext) {
        if ($env:IDENTITY_ENDPOINT -or $env:MSI_ENDPOINT -or (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
            Write-Log 'Connecting to Azure with managed identity to request an Exchange token...' | Out-Host
            Disable-AzContextAutosave -Scope Process -ErrorAction SilentlyContinue | Out-Null
            Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        }
        else {
            throw 'No Azure context found. Run Connect-AzAccount first for local testing.'
        }
    }

    $token = (Get-AzAccessToken -ResourceUrl 'https://outlook.office365.com' -ErrorAction Stop).Token
    if ($token -is [securestring]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
        try {
            $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            if ($bstr -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
    }

    Write-Log 'Exchange Online access token acquired from Az context.' -Level SUCCESS | Out-Host
    return [string]$token
}

function Get-TargetMailbox {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    try {
        return Get-Mailbox -Identity $Identity -ErrorAction Stop
    }
    catch {
        Write-Log "Get-Mailbox could not resolve [$Identity] directly. Trying Exchange recipient fallback..." -Level WARN | Out-Host
    }

    try {
        $exoRecipient = Get-EXORecipient -Identity $Identity -Properties RecipientTypeDetails,PrimarySmtpAddress -ErrorAction Stop
        if ($exoRecipient -and $exoRecipient.RecipientTypeDetails -in @('UserMailbox','SharedMailbox')) {
            return Get-Mailbox -Identity $exoRecipient.PrimarySmtpAddress.ToString() -ErrorAction Stop
        }
    }
    catch {}

    try {
        $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
        if ($recipient -and $recipient.RecipientTypeDetails -in @('UserMailbox','SharedMailbox')) {
            return Get-Mailbox -Identity $recipient.Identity -ErrorAction Stop
        }
    }
    catch {}

    return $null
}

try {
    Write-Log "Starting mailbox conversion for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No mailbox changes will be made.' -Level WARN
    }

    Write-Log 'Importing ExchangeOnlineManagement module...'
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connected = $false
    Write-Log 'Connecting to Exchange Online...'

    try {
        $exoToken = Get-ExchangeAccessToken
        Connect-ExchangeOnline -AccessToken $exoToken -Organization $ExchangeOrganization -ShowBanner:$false -ErrorAction Stop
        $connected = $true
        Write-Log 'Connected to Exchange Online using access token.' -Level SUCCESS
    }
    catch {
        Write-Log "Access-token connection failed. Falling back to managed identity. Details: $($_.Exception.Message)" -Level WARN
    }

    if (-not $connected) {
        Connect-ExchangeOnline -ManagedIdentity -Organization $ExchangeOrganization -ShowBanner:$false -ErrorAction Stop
        Write-Log 'Connected to Exchange Online using managed identity.' -Level SUCCESS
    }

    $mailbox = Get-TargetMailbox -Identity $UserPrincipalName
    if ($null -eq $mailbox) {
        Write-Log "No Exchange mailbox was found for [$UserPrincipalName]. Skipping mailbox conversion." -Level WARN
        return
    }

    Write-Log "Mailbox found. Current type: [$($mailbox.RecipientTypeDetails)]"

    if ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') {
        Write-Log 'Mailbox is already a SharedMailbox. No change is required.' -Level SUCCESS
        return
    }

    if ($simulateOnly) {
        Write-Log '[DryRun] Would convert mailbox to SharedMailbox.' -Level WARN
        return
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Convert mailbox to SharedMailbox')) {
        Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop

        $verificationDeadline = (Get-Date).AddMinutes(3)
        $verifiedMailbox = $null
        do {
            Start-Sleep -Seconds 10
            $verifiedMailbox = Get-TargetMailbox -Identity $UserPrincipalName
            if ($null -eq $verifiedMailbox) {
                throw "Mailbox verification could not find [$UserPrincipalName] after conversion request."
            }
            Write-Log "Verification check - current type: [$($verifiedMailbox.RecipientTypeDetails)]"
            if ($verifiedMailbox.RecipientTypeDetails -eq 'SharedMailbox') {
                break
            }
        } while ((Get-Date) -lt $verificationDeadline)

        if ($null -eq $verifiedMailbox -or $verifiedMailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            throw "Mailbox conversion verification failed after waiting for propagation. Current type is [$($verifiedMailbox.RecipientTypeDetails)]."
        }

        Write-Log 'Mailbox converted to SharedMailbox successfully.' -Level SUCCESS
    }

    Write-Log 'Step 02 completed successfully.' -Level SUCCESS
    return
}
catch {
    Write-Log "Step 02 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
    catch {}
}
