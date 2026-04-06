<#
.SYNOPSIS
    Step 01 | Immediate Access Lockdown
.DESCRIPTION
    First-step runbook for offboarding.

    This script performs only the immediate access-blocking actions that must
    happen before account disablement:
      1. Revoke all current sign-in sessions
      2. Reset the password to a random complex value
      3. Remove registered authentication methods (MFA / Authenticator / phone /
         FIDO2 / Windows Hello for Business / software OATH / email / TAP)

    The account is NOT disabled here.
.PARAMETER UserPrincipalName
    User to secure immediately.
.PARAMETER LogPath
    Optional log file path.
.NOTES
    Microsoft Graph permissions required:
      - User.ReadWrite.All
      - User-PasswordProfile.ReadWrite.All
      - User.RevokeSessions.All
      - UserAuthenticationMethod.ReadWrite.All
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-01_$(Get-Date -Format 'yyyy-MM-dd').log",

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
    $entry = "[$timestamp] [$Level] [STEP-01] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-01] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
}

function New-RandomPassword {
    $upper   = [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lower   = [char[]]'abcdefghijklmnopqrstuvwxyz'
    $digits  = [char[]]'0123456789'
    $special = [char[]]'!@#$%^&*()-_=+[]{}|;:,.<>?'
    $all     = $upper + $lower + $digits + $special
    $rng     = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    function Get-SecureChar {
        param([char[]]$CharacterSet)
        $buffer = [byte[]]::new(1)
        $rng.GetBytes($buffer)
        return $CharacterSet[$buffer[0] % $CharacterSet.Count]
    }

    $passwordChars = @(
        (Get-SecureChar -CharacterSet $upper)
        (Get-SecureChar -CharacterSet $lower)
        (Get-SecureChar -CharacterSet $digits)
        (Get-SecureChar -CharacterSet $special)
    )

    1..20 | ForEach-Object {
        $buffer = [byte[]]::new(1)
        $rng.GetBytes($buffer)
        $passwordChars += $all[$buffer[0] % $all.Count]
    }

    return -join ($passwordChars | Sort-Object { Get-Random })
}

function Get-GraphAccessToken {
    if (-not (Get-Command -Name Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        throw 'Az.Accounts is required. Install/import Az.Accounts and sign in with Connect-AzAccount, or run inside Azure Automation with managed identity.'
    }

    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $azContext) {
        if ($env:IDENTITY_ENDPOINT -or $env:MSI_ENDPOINT -or (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
            Write-Log 'Connecting to Azure with managed identity to request a Graph token...' | Out-Host
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
        [ValidateSet('GET','POST','PATCH','DELETE')]
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

    if ($Method -in @('POST','PATCH')) {
        $requestParams['ContentType'] = 'application/json'
    }

    if ($null -ne $Body) {
        $requestParams['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }
    elseif ($Method -eq 'POST') {
        $requestParams['Body'] = '{}'
    }

    return Invoke-RestMethod @requestParams
}

function Remove-AuthenticationMethods {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$UserUpn,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [bool]$UseWhatIf = $false
    )

    $removed = 0
    $skipped = 0

    Write-Log 'Retrieving authentication methods...' | Out-Host
    $methodsResponse = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/authentication/methods" -AccessToken $AccessToken
    $methods = @($methodsResponse.value)

    if ($methods.Count -eq 0) {
        Write-Log 'No authentication methods returned by Graph.' -Level WARN | Out-Host
        return [PSCustomObject]@{ Removed = 0; Skipped = 0 }
    }

    foreach ($method in $methods) {
        $methodId = $method.id
        $odataType = $method.'@odata.type'
        $targetUri = $null
        $label = $odataType

        switch ($odataType) {
            '#microsoft.graph.passwordAuthenticationMethod' {
                Write-Log 'Skipping password method (cannot be deleted).' -Level INFO | Out-Host
                $skipped++
                continue
            }
            '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/microsoftAuthenticatorMethods/$methodId"
                $label = 'Microsoft Authenticator'
            }
            '#microsoft.graph.phoneAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/phoneMethods/$methodId"
                $label = 'Phone / SMS / Voice'
            }
            '#microsoft.graph.fido2AuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/fido2Methods/$methodId"
                $label = 'FIDO2 Security Key'
            }
            '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/windowsHelloForBusinessMethods/$methodId"
                $label = 'Windows Hello for Business'
            }
            '#microsoft.graph.softwareOathAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/softwareOathMethods/$methodId"
                $label = 'Software OATH Token'
            }
            '#microsoft.graph.emailAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/emailMethods/$methodId"
                $label = 'Email Authentication Method'
            }
            '#microsoft.graph.temporaryAccessPassAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/temporaryAccessPassMethods/$methodId"
                $label = 'Temporary Access Pass'
            }
            '#microsoft.graph.platformCredentialAuthenticationMethod' {
                $targetUri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/platformCredentialMethods/$methodId"
                $label = 'Platform Credential / Passkey'
            }
            default {
                Write-Log "Unsupported or non-removable authentication method detected: [$odataType]" -Level WARN | Out-Host
                $skipped++
                continue
            }
        }

        try {
            if ($UseWhatIf) {
                Write-Log "[WhatIf] Would remove authentication method: [$label]" -Level WARN | Out-Host
                continue
            }

            Invoke-GraphRequest -Method DELETE -Uri $targetUri -AccessToken $AccessToken | Out-Null
            Write-Log "Removed authentication method: [$label]" -Level SUCCESS | Out-Host
            $removed++
        }
        catch {
            Write-Log "Failed to remove authentication method [$label]: $($_.Exception.Message)" -Level WARN | Out-Host
            $skipped++
        }
    }

    return [PSCustomObject]@{
        Removed = $removed
        Skipped = $skipped
    }
}

try {
    Write-Log "Starting immediate access lockdown for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No changes will be made.' -Level WARN
    }

    $graphToken = Get-GraphAccessToken
    $userLookupUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=id,displayName,accountEnabled,signInSessionsValidFromDateTime"
    $user = Invoke-GraphRequest -Method GET -Uri $userLookupUri -AccessToken $graphToken
    Write-Log "User found: [$($user.displayName)] | Enabled: [$($user.accountEnabled)]"

    if ($simulateOnly) {
        Write-Log '[DryRun] Would revoke all active sign-in sessions.' -Level WARN
    }
    elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Revoke all active sign-in sessions')) {
        $revokeResult = Invoke-GraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)/revokeSignInSessions" -AccessToken $graphToken
        if ($revokeResult.value -eq $true) {
            Write-Log 'Session revocation accepted by Microsoft Graph.' -Level SUCCESS
        }
        else {
            Write-Log 'Session revocation returned a non-standard response, but the request completed.' -Level WARN
        }
    }

    if ($simulateOnly) {
        Write-Log '[DryRun] Would reset the password and force password change at next sign-in.' -Level WARN
    }
    elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Reset password and force password change')) {
        $newPassword = New-RandomPassword
        $body = @{ passwordProfile = @{ password = $newPassword; forceChangePasswordNextSignIn = $true } }

        try {
            Invoke-GraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)" -AccessToken $graphToken -Body $body | Out-Null
            Write-Log 'Password reset completed and force-change-next-sign-in is enabled.' -Level SUCCESS
        }
        finally {
            $newPassword = $null
            [System.GC]::Collect()
        }
    }

    $mfaResult = Remove-AuthenticationMethods -UserId $user.id -UserUpn $UserPrincipalName -AccessToken $graphToken -UseWhatIf:$simulateOnly
    Write-Log "Authentication method cleanup finished. Removed: [$($mfaResult.Removed)] | Skipped: [$($mfaResult.Skipped)]"

    Write-Log 'Step 01 completed successfully.' -Level SUCCESS
    exit 0
}
catch {
    Write-Log "Step 01 failed: $($_.Exception.Message)" -Level ERROR
    exit 1
}
