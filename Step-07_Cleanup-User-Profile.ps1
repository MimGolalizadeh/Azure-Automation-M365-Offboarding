<#
.SYNOPSIS
    Step 07 | Cleanup User Profile and Remove Manager
.DESCRIPTION
    Cleans up non-essential Entra ID / Microsoft 365 profile attributes for an
    offboarded user and removes the manager relationship.

    This step intentionally preserves the core identity values needed for audit
    and mailbox continuity, while clearing day-to-day profile data that is no
    longer needed after offboarding.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Simulate changes without modifying the user.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-07_$(Get-Date -Format 'yyyy-MM-dd').log",

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
    $entry = "[$timestamp] [$Level] [STEP-07] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-07] Could not write log file [$LogPath]: $($_.Exception.Message)"
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
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $false)][object]$Body
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }

    if ($PSBoundParameters.ContainsKey('Body')) {
        $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $jsonBody -ContentType 'application/json' -ErrorAction Stop
    }

    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
}

function Get-StatusCodeFromException {
    param([Parameter(Mandatory = $true)][object]$ExceptionObject)

    try {
        if ($ExceptionObject.Response -and $ExceptionObject.Response.StatusCode) {
            return [int]$ExceptionObject.Response.StatusCode
        }
    }
    catch {}

    return $null
}

try {
    Write-Log "Starting user-profile cleanup for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No Entra profile changes will be made.' -Level WARN
    }

    $graphToken = Get-GraphAccessToken
    $encodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)

    $identityUri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}?`$select=id,displayName,givenName,surname,userPrincipalName,userType,createdDateTime,companyName,employeeId,employeeHireDate,department,jobTitle,officeLocation,city,state,country,businessPhones,mobilePhone,otherMails"
    $identitySnapshot = Invoke-GraphRequest -Method GET -Uri $identityUri -AccessToken $graphToken

    Write-Log "User found: [$($identitySnapshot.displayName)] | UPN: [$($identitySnapshot.userPrincipalName)] | Type: [$($identitySnapshot.userType)]"

    $managerExists = $false
    try {
        $manager = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUpn/manager?`$select=id,displayName,userPrincipalName" -AccessToken $graphToken
        $managerExists = $true
        Write-Log "Current manager: [$($manager.displayName)] | UPN: [$($manager.userPrincipalName)]"
    }
    catch {
        $statusCode = Get-StatusCodeFromException -ExceptionObject $_.Exception
        if ($statusCode -eq 404) {
            Write-Log 'No manager relationship is currently set.' -Level INFO
        }
        else {
            Write-Log "Could not read manager relationship: $($_.Exception.Message)" -Level WARN
        }
    }

    $platformRestrictedFields = @(
        'aboutMe', 'interests', 'pastProjects', 'responsibilities', 'schools', 'skills'
    )

    $nonNullableFields = @(
        'birthday', 'usageLocation'
    )

    $cleanupFields = [ordered]@{
        businessPhones    = @()
        city              = $null
        country           = $null
        department        = $null
        employeeType      = $null
        faxNumber         = $null
        jobTitle          = $null
        mobilePhone       = $null
        officeLocation    = $null
        otherMails        = @()
        postalCode        = $null
        preferredLanguage = $null
        state             = $null
        streetAddress     = $null
    }

    if ($simulateOnly) {
        if ($managerExists) {
            Write-Log '[DryRun] Would remove the manager relationship.' -Level WARN
        }
        else {
            Write-Log '[DryRun] Manager relationship is already empty.' -Level INFO
        }

        Write-Log "[DryRun] Would clear profile fields: $($cleanupFields.Keys -join ', ')" -Level WARN
        if ($platformRestrictedFields.Count -gt 0) {
            Write-Log "Skipped (platform-restricted in app-only context): $($platformRestrictedFields -join ', ')" -Level INFO
        }
        if ($nonNullableFields.Count -gt 0) {
            Write-Log "Skipped (non-nullable/system-managed): $($nonNullableFields -join ', ')" -Level INFO
        }

        Write-Log 'Step 07 dry-run completed successfully.' -Level SUCCESS
        return
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Remove Entra manager relationship')) {
        try {
            Invoke-GraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$encodedUpn/manager/`$ref" -AccessToken $graphToken | Out-Null
            Write-Log 'Manager relationship removed.' -Level SUCCESS
        }
        catch {
            $statusCode = Get-StatusCodeFromException -ExceptionObject $_.Exception
            if ($statusCode -eq 404) {
                Write-Log 'Manager relationship is already empty. Nothing to remove.' -Level INFO
            }
            else {
                Write-Log "Failed to remove manager relationship: $($_.Exception.Message)" -Level WARN
            }
        }
    }

    if ($platformRestrictedFields.Count -gt 0) {
        Write-Log "Skipped (platform-restricted in app-only context): $($platformRestrictedFields -join ', ')" -Level INFO
    }
    if ($nonNullableFields.Count -gt 0) {
        Write-Log "Skipped (non-nullable/system-managed): $($nonNullableFields -join ', ')" -Level INFO
    }

    $clearedFields = @()
    $failedFields = @()
    $patchUri = "https://graph.microsoft.com/v1.0/users/$encodedUpn"

    foreach ($field in $cleanupFields.Keys) {
        try {
            $body = @{ $field = $cleanupFields[$field] }
            Invoke-GraphRequest -Method PATCH -Uri $patchUri -AccessToken $graphToken -Body $body | Out-Null
            $clearedFields += $field
        }
        catch {
            $failedFields += $field
            Write-Log "Could not clear [$field]: $($_.Exception.Message)" -Level WARN
        }
    }

    Write-Log "Profile cleanup summary - cleared: [$($clearedFields.Count)] | failed: [$($failedFields.Count)] | skipped: [$($platformRestrictedFields.Count + $nonNullableFields.Count)]"
    if ($clearedFields.Count -gt 0) {
        Write-Log "Cleared fields: $($clearedFields -join ', ')" -Level SUCCESS
    }
    if ($failedFields.Count -gt 0) {
        Write-Log "Fields not cleared: $($failedFields -join ', ')" -Level WARN
    }

    try {
        $verifyFields = @('businessPhones','city','country','department','employeeType','faxNumber','jobTitle','mobilePhone','officeLocation','otherMails','postalCode','preferredLanguage','state','streetAddress')
        $verify = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encodedUpn}?`$select=$($verifyFields -join ',')" -AccessToken $graphToken

        $remaining = @()
        foreach ($field in $verifyFields) {
            $value = $verify.$field
            if ($null -eq $value) { continue }

            if ($value -is [System.Array]) {
                if ($value.Count -gt 0) { $remaining += $field }
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                $remaining += $field
            }
        }

        try {
            Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUpn/manager?`$select=id" -AccessToken $graphToken | Out-Null
            $remaining += 'manager'
        }
        catch {
            $statusCode = Get-StatusCodeFromException -ExceptionObject $_.Exception
            if ($statusCode -ne 404) {
                Write-Log "Could not fully verify manager removal: $($_.Exception.Message)" -Level WARN
            }
        }

        if ($remaining.Count -eq 0) {
            Write-Log 'Verification passed: manager relationship and target profile fields are cleared.' -Level SUCCESS
        }
        else {
            Write-Log "Verification warning: some values still remain: $($remaining -join ', ')" -Level WARN
        }
    }
    catch {
        Write-Log "Could not complete post-cleanup verification: $($_.Exception.Message)" -Level WARN
    }

    Write-Log 'Step 07 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 07 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
