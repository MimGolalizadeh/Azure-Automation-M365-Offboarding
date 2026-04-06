<#
.SYNOPSIS
    Step 03 | Remove User Group Memberships
.DESCRIPTION
    Removes the user from Entra ID / Microsoft 365 groups while preserving the
    designated offboarding group.

    This step is intentionally limited to group membership cleanup only.
    It does not remove licenses or distribution lists.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER OffboardingGroupName
    Display name or object ID of the offboarding group to keep.
    If omitted, the script attempts to read it from Azure Automation variables.
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Simulate removals without changing memberships.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$OffboardingGroupName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-03_$(Get-Date -Format 'yyyy-MM-dd').log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OffboardingGroupName) -and (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
    foreach ($name in @('OffboardingGroupName', 'OFFBOARDINGGROUPNAME', 'OffboardingGroup', 'OffboardingGroupId')) {
        try {
            $val = Get-AutomationVariable -Name $name -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $OffboardingGroupName = "$val"
                break
            }
        }
        catch {}
    }
}

$OffboardingGroupName = "$OffboardingGroupName".Trim().Trim('"').Trim("'")
if ([string]::IsNullOrWhiteSpace($OffboardingGroupName)) {
    throw "OffboardingGroupName is required. Pass -OffboardingGroupName or create an Azure Automation variable such as 'OffboardingGroupName'."
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [STEP-03] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-03] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
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

function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    $items = @()
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $response = Invoke-GraphRequest -Method GET -Uri $next -AccessToken $AccessToken
        if ($response.value) {
            $items += @($response.value)
        }
        $next = $response.'@odata.nextLink'
    }

    return $items
}

function Resolve-OffboardingGroup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    if ($Value -match '^[0-9a-fA-F-]{36}$') {
        return Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$Value?`$select=id,displayName" -AccessToken $AccessToken
    }

    $safeValue = $Value.Replace("'", "''")
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safeValue'&`$select=id,displayName"
    $response = Invoke-GraphRequest -Method GET -Uri $uri -AccessToken $AccessToken

    if (-not $response.value -or $response.value.Count -eq 0) {
        throw "Offboarding group [$Value] was not found in Microsoft Graph."
    }

    return $response.value[0]
}

function Test-DynamicGroup {
    param([object]$Group)

    if ($Group.membershipRule) { return $true }
    if ($Group.groupTypes -and ($Group.groupTypes -contains 'DynamicMembership')) { return $true }
    return $false
}

function Test-ExchangeManagedGroup {
    param([object]$Group)

    $groupTypes = @($Group.groupTypes)
    $isUnified = $groupTypes -contains 'Unified'
    return ([bool]$Group.mailEnabled -and -not $isUnified)
}

function Test-RoleAssignableGroup {
    param([object]$Group)

    return [bool]$Group.isAssignableToRole
}

try {
    Write-Log "Starting group removal step for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No memberships will be changed.' -Level WARN
    }

    $graphToken = Get-GraphAccessToken

    $userLookupUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=id,displayName,userPrincipalName"
    $user = Invoke-GraphRequest -Method GET -Uri $userLookupUri -AccessToken $graphToken
    Write-Log "User found: [$($user.displayName)] | UPN: [$($user.userPrincipalName)]"

    $protectedGroup = Resolve-OffboardingGroup -Value $OffboardingGroupName -AccessToken $graphToken
    Write-Log "Protected offboarding group resolved: [$($protectedGroup.displayName)] | ID: [$($protectedGroup.id)]"

    $memberOfUri = "https://graph.microsoft.com/v1.0/users/$($user.id)/memberOf?`$top=999&`$select=id,displayName,groupTypes,membershipRule,mailEnabled,securityEnabled,isAssignableToRole"
    $memberships = Get-GraphCollection -Uri $memberOfUri -AccessToken $graphToken
    $groups = @($memberships | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' })

    Write-Log "Found [$($groups.Count)] group membership(s) to evaluate."

    $removedCount = 0
    $skippedCount = 0

    foreach ($group in $groups) {
        $groupId = $group.id
        $groupName = if ($group.displayName) { $group.displayName } else { $groupId }

        if ($groupId -eq $protectedGroup.id -or $groupName -eq $protectedGroup.displayName) {
            Write-Log "SKIP protected offboarding group: [$groupName]" -Level WARN
            $skippedCount++
            continue
        }

        if (Test-DynamicGroup -Group $group) {
            Write-Log "SKIP dynamic group: [$groupName]" -Level WARN
            $skippedCount++
            continue
        }

        if (Test-ExchangeManagedGroup -Group $group) {
            Write-Log "SKIP Exchange-managed mail-enabled group (handled in Step 04): [$groupName]" -Level WARN
            $skippedCount++
            continue
        }

        if (Test-RoleAssignableGroup -Group $group) {
            Write-Log "Group [$groupName] is role-assignable; removal may require elevated Graph role-management permissions." -Level INFO
        }

        try {
            if ($simulateOnly) {
                Write-Log "[DryRun] Would remove user from group: [$groupName]" -Level WARN
                continue
            }

            if ($PSCmdlet.ShouldProcess($groupName, 'Remove user from group')) {
                $removeUri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/$($user.id)/`$ref"
                Invoke-GraphRequest -Method DELETE -Uri $removeUri -AccessToken $graphToken | Out-Null
                Write-Log "Removed user from group: [$groupName]" -Level SUCCESS
                $removedCount++
            }
        }
        catch {
            $message = $_.Exception.Message
            if ((Test-RoleAssignableGroup -Group $group) -and $message -match '403|Forbidden') {
                Write-Log "Failed to remove role-assignable group [$groupName]. The Automation managed identity likely needs RoleManagement.ReadWrite.Directory and the Privileged Role Administrator role. Details: $message" -Level WARN
            }
            else {
                Write-Log "Failed to remove user from [$groupName]: $message" -Level WARN
            }
            $skippedCount++
        }
    }

    Write-Log "Group removal finished. Removed: [$removedCount] | Skipped: [$skippedCount]"

    $remainingMemberships = Get-GraphCollection -Uri $memberOfUri -AccessToken $graphToken
    $remainingGroups = @($remainingMemberships | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' })
    $remainingNonExcludedGroups = @($remainingGroups | Where-Object {
        $_.id -ne $protectedGroup.id -and
        -not (Test-DynamicGroup -Group $_) -and
        -not (Test-ExchangeManagedGroup -Group $_)
    })

    if ($remainingNonExcludedGroups.Count -gt 0 -and -not $simulateOnly) {
        $remainingNames = ($remainingNonExcludedGroups | ForEach-Object { $_.displayName }) -join ', '
        throw "Verification failed. User is still a member of non-excluded group(s): $remainingNames"
    }

    if ($remainingNonExcludedGroups.Count -eq 0) {
        Write-Log 'Verification passed. No remaining removable groups were found.' -Level SUCCESS
    }
    else {
        $remainingNames = ($remainingNonExcludedGroups | ForEach-Object { $_.displayName }) -join ', '
        Write-Log "[DryRun] Remaining removable groups observed: $remainingNames" -Level WARN
    }

    Write-Log 'Step 03 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 03 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
