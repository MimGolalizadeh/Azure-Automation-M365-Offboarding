<#
.SYNOPSIS
    Step 08 | Remove Remaining User Licenses
.DESCRIPTION
    Removes all remaining Microsoft 365 licenses from the target user.

    This step is intentionally focused on license cleanup only. Licenses still
    inherited through group-based licensing are reported clearly, so any delay
    from group-membership propagation is visible in the job output.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Simulate changes without removing licenses.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-08_$(Get-Date -Format 'yyyy-MM-dd').log",

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
    $entry = "[$timestamp] [$Level] [STEP-08] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-08] Could not write log file [$LogPath]: $($_.Exception.Message)"
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

function Get-LicenseDetails {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $response = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/licenseDetails?`$select=skuId,skuPartNumber" -AccessToken $AccessToken
    if ($null -eq $response) { return @() }
    if ($response.value) { return @($response.value) }
    return @($response)
}

try {
    Write-Log "Starting license cleanup for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No license changes will be made.' -Level WARN
    }

    $graphToken = Get-GraphAccessToken
    $encodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)

    $user = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encodedUpn}?`$select=id,displayName,userPrincipalName,assignedLicenses,licenseAssignmentStates" -AccessToken $graphToken
    Write-Log "User found: [$($user.displayName)] | UPN: [$($user.userPrincipalName)]"

    $licenseDetails = @(Get-LicenseDetails -UserId $user.id -AccessToken $graphToken)
    if ($licenseDetails.Count -eq 0) {
        Write-Log 'No remaining licenses are assigned. Nothing to remove.' -Level SUCCESS
        return
    }

    $licenseStatesBySku = @{}
    foreach ($state in @($user.licenseAssignmentStates)) {
        $skuKey = [string]$state.skuId
        if (-not $licenseStatesBySku.ContainsKey($skuKey)) {
            $licenseStatesBySku[$skuKey] = @()
        }
        $licenseStatesBySku[$skuKey] += $state
    }

    Write-Log "Found [$($licenseDetails.Count)] remaining license(s): $((@($licenseDetails | ForEach-Object { $_.skuPartNumber }) | Where-Object { $_ } ) -join ', ')"

    $removedCount = 0
    $skippedCount = 0

    foreach ($license in $licenseDetails) {
        $skuId = [string]$license.skuId
        $skuName = if ([string]::IsNullOrWhiteSpace([string]$license.skuPartNumber)) { $skuId } else { [string]$license.skuPartNumber }

        $states = @()
        if ($licenseStatesBySku.ContainsKey($skuId)) {
            $states = @($licenseStatesBySku[$skuId])
        }

        $hasDirectAssignment = @($states | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.assignedByGroup) }).Count -gt 0
        $hasGroupAssignment = @($states | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.assignedByGroup) }).Count -gt 0
        $source = if ($hasDirectAssignment -and $hasGroupAssignment) { 'Direct + Group' } elseif ($hasDirectAssignment) { 'Direct' } elseif ($hasGroupAssignment) { 'Group' } else { 'Unknown' }

        if ($simulateOnly) {
            Write-Log "[DryRun] Would remove license [$skuName] | Source: [$source]" -Level WARN
            continue
        }

        if (-not $PSCmdlet.ShouldProcess($UserPrincipalName, "Remove license [$skuName]")) {
            continue
        }

        try {
            Invoke-GraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)/assignLicense" -AccessToken $graphToken -Body @{
                addLicenses    = @()
                removeLicenses = @($skuId)
            } | Out-Null

            Write-Log "Removed license [$skuName]" -Level SUCCESS
            $removedCount++
        }
        catch {
            $message = $_.Exception.Message
            if ($message -match 'cannotRemoveLicenseAssignedViaGroup|inherited from a group') {
                Write-Log "Skipped group-based license [$skuName]; it should clear after group-license propagation." -Level WARN
            }
            else {
                Write-Log "Could not remove license [$skuName]: $message" -Level WARN
            }
            $skippedCount++
        }
    }

    $remainingLicenses = @(Get-LicenseDetails -UserId $user.id -AccessToken $graphToken)
    Write-Log "License removal summary - removed: [$removedCount] | skipped/warnings: [$skippedCount] | remaining now: [$($remainingLicenses.Count)]"

    if ($remainingLicenses.Count -eq 0) {
        Write-Log 'Verification passed: no remaining licenses detected.' -Level SUCCESS
    }
    else {
        Write-Log "Remaining license(s): $((@($remainingLicenses | ForEach-Object { $_.skuPartNumber }) | Where-Object { $_ }) -join ', ')" -Level WARN
    }

    Write-Log 'Step 08 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 08 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
