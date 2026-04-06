<#
.SYNOPSIS
    Step 09 | Disable User Account
.DESCRIPTION
    Final offboarding step to disable the user account in Microsoft Entra ID.

    This script is intentionally focused on account disablement only. It reads
    the user's current state from Microsoft Graph, disables the account when it
    is still enabled, and verifies the final result.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Simulate the change without disabling the account.
.NOTES
    Microsoft Graph application permissions typically required:
      - User.Read.All
      - User.EnableDisableAccount.All (recommended for accountEnabled)
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-09_$(Get-Date -Format 'yyyy-MM-dd').log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [STEP-09] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-09] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
}

function Get-GraphAccessToken {
    if (-not (Get-Command -Name Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        throw 'Az.Accounts is required. Install/import Az.Accounts and sign in with Connect-AzAccount, or run inside Azure Automation with managed identity.'
    }

    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $azContext) {
        if ($env:IDENTITY_ENDPOINT -or $env:MSI_ENDPOINT -or (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
            Write-Log 'Connecting to Azure with managed identity to request a Microsoft Graph token...' | Out-Host
            Disable-AzContextAutosave -Scope Process -ErrorAction SilentlyContinue | Out-Null
            Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        }
        else {
            throw 'No Azure context found. Run Connect-AzAccount first for local testing.'
        }
    }

    $token = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).Token
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

    Write-Log 'Microsoft Graph access token acquired from Az context.' -Level SUCCESS | Out-Host
    return [string]$token
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET','PATCH')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [object]$Body
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    $requestParams = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }

    if ($Method -eq 'PATCH') {
        $requestParams['ContentType'] = 'application/json'
    }

    if ($null -ne $Body) {
        $requestParams['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }

    return Invoke-RestMethod @requestParams
}

try {
    Write-Log "Starting account disable step for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No account changes will be made.' -Level WARN
    }

    $graphToken = Get-GraphAccessToken
    $encodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)

    $user = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encodedUpn}?`$select=id,displayName,userPrincipalName,accountEnabled,userType" -AccessToken $graphToken
    Write-Log "User found: [$($user.displayName)] | UPN: [$($user.userPrincipalName)] | AccountEnabled: [$($user.accountEnabled)] | Type: [$($user.userType)]"

    if ($user.accountEnabled -eq $false) {
        Write-Log 'Account is already disabled. No change required.' -Level SUCCESS
        return
    }

    if ($simulateOnly) {
        Write-Log "[DryRun] Would disable account for [$($user.userPrincipalName)]" -Level WARN
        return
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Disable Microsoft Entra account')) {
        Invoke-GraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)" -AccessToken $graphToken -Body @{
            accountEnabled = $false
        } | Out-Null

        Write-Log "Account disable request sent for [$($user.userPrincipalName)]" -Level SUCCESS
    }

    $verify = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)?`$select=accountEnabled" -AccessToken $graphToken
    if ($verify.accountEnabled -eq $false) {
        Write-Log 'Verification passed: AccountEnabled is now [False].' -Level SUCCESS
    }
    else {
        Write-Log "Verification warning: AccountEnabled still reports [$($verify.accountEnabled)]." -Level WARN
    }

    Write-Log 'Step 09 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 09 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
