<#
.SYNOPSIS
    Step 06 | Remove Calendar Meetings, Events, and Room Booking Access
.DESCRIPTION
    Performs Exchange-side calendar cleanup for an offboarded user by:
      - cancelling future meetings organized by the user
      - removing explicit sharing/delegate permissions from the user's Calendar
      - removing room/resource booking and approval access where the user is listed

    This step is intentionally Exchange-focused because Exchange cmdlets are
    generally more reliable than Microsoft Graph for mailbox-side cleanup in
    app-only automation scenarios.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER ExchangeOrganization
    Exchange Online organization value, for example: contoso.onmicrosoft.com
.PARAMETER DaysAhead
    Number of days ahead to process for organized meetings.
.PARAMETER LogPath
    Optional log path.
.PARAMETER DryRun
    Simulate changes without removing anything.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$ExchangeOrganization,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$DaysAhead = 365,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-06_$(Get-Date -Format 'yyyy-MM-dd').log",

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
    $entry = "[$timestamp] [$Level] [STEP-06] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-06] Could not write log file [$LogPath]: $($_.Exception.Message)"
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

function Get-GraphAccessToken {
    if (-not (Get-Command -Name Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        throw 'Az.Accounts is required to request a Microsoft Graph token.'
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

function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken
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

    return @($items)
}

function Get-SafeArray {
    param([object]$Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

function Test-TokenMatch {
    param(
        [AllowNull()][AllowEmptyString()][string]$Value,
        [string[]]$Tokens
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $normalizedValue = $Value.ToLower()

    foreach ($token in @($Tokens | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $normalizedToken = $token.ToLower()
        if ($normalizedValue -eq $normalizedToken -or $normalizedValue -like "*$normalizedToken*") {
            return $true
        }
    }

    return $false
}

function Split-MatchedValues {
    param(
        [object]$Values,
        [string[]]$Tokens
    )

    $kept = New-Object System.Collections.Generic.List[string]
    $removed = New-Object System.Collections.Generic.List[string]

    foreach ($value in (Get-SafeArray -Value $Values)) {
        $text = [string]$value
        if (Test-TokenMatch -Value $text -Tokens $Tokens) {
            [void]$removed.Add($text)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$kept.Add($text)
        }
    }

    return [pscustomobject]@{
        Kept    = @($kept)
        Removed = @($removed)
    }
}

try {
    Write-Log "Starting calendar and room cleanup for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No Exchange changes will be made.' -Level WARN
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

    $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    $recipient = Get-Recipient -Identity $UserPrincipalName -ErrorAction SilentlyContinue

    $userEmail = [string]$mailbox.PrimarySmtpAddress
    $userAlias = [string]$mailbox.Alias
    $userDisplay = [string]$mailbox.DisplayName
    $userDn = if ($recipient) { [string]$recipient.DistinguishedName } else { $null }
    $userExternalId = if ($recipient) { [string]$recipient.ExternalDirectoryObjectId } else { $null }

    $matchTokens = @(
        $UserPrincipalName,
        $UserPrincipalName.ToLower(),
        $userEmail,
        $userAlias,
        $userDisplay,
        $userDn,
        $userExternalId
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    Write-Log "Mailbox found: [$userDisplay] | Email: [$userEmail] | Type: [$($mailbox.RecipientTypeDetails)]"

    $meetingActions = 0
    $calendarPermissionsRemoved = 0
    $roomAssignmentsRemoved = 0
    $skipped = 0

    Write-Log "Checking future organized meetings for the next [$DaysAhead] day(s)..."
    try {
        $preview = @(Remove-CalendarEvents -Identity $userEmail -CancelOrganizedMeetings -QueryStartDate (Get-Date) -QueryWindowInDays $DaysAhead -PreviewOnly -UseCustomRouting -ErrorAction Stop)
        $previewCount = @($preview).Count
        $meetingActions = $previewCount

        if ($simulateOnly) {
            Write-Log "[DryRun] Preview identified [$previewCount] future organized meeting item(s) to cancel/release." -Level WARN
        }
        elseif ($PSCmdlet.ShouldProcess($userEmail, "Cancel future organized meetings and room reservations")) {
            Remove-CalendarEvents -Identity $userEmail -CancelOrganizedMeetings -QueryStartDate (Get-Date) -QueryWindowInDays $DaysAhead -UseCustomRouting -Confirm:$false -ErrorAction Stop | Out-Null
            Write-Log "Cancelled future organized meetings for [$userEmail] across the next [$DaysAhead] day(s)." -Level SUCCESS
        }
    }
    catch {
        Write-Log "Exchange meeting cancellation failed: $($_.Exception.Message)" -Level WARN
        Write-Log 'Trying Microsoft Graph fallback for organized meetings...' -Level WARN

        try {
            $graphToken = Get-GraphAccessToken
            $graphUser = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=id,userPrincipalName" -AccessToken $graphToken
            $startIso = (Get-Date).ToString('o')
            $endIso = (Get-Date).AddDays($DaysAhead).ToString('o')
            $calendarViewUri = "https://graph.microsoft.com/v1.0/users/$($graphUser.id)/calendarView?startDateTime=$([uri]::EscapeDataString($startIso))&endDateTime=$([uri]::EscapeDataString($endIso))&`$select=id,subject,isOrganizer,isCancelled,start"
            $organizedMeetings = @(Get-GraphCollection -Uri $calendarViewUri -AccessToken $graphToken | Where-Object {
                $_.isOrganizer -eq $true -and $_.isCancelled -ne $true
            })
            $meetingActions = @($organizedMeetings).Count

            if ($simulateOnly) {
                Write-Log "[DryRun][GraphFallback] Would cancel [$meetingActions] organized meeting item(s)." -Level WARN
            }
            else {
                $organizedMeetings | ForEach-Object {
                    $subject = if ([string]::IsNullOrWhiteSpace([string]$_.subject)) { '(no subject)' } else { [string]$_.subject }
                    if ($PSCmdlet.ShouldProcess($subject, 'Cancel meeting using Microsoft Graph fallback')) {
                        Invoke-GraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($graphUser.id)/events/$($_.id)/cancel" -AccessToken $graphToken -Body @{ Comment = 'This event has been cancelled because the organizer is no longer with the organization.' } | Out-Null
                        Write-Log "Cancelled meeting using Microsoft Graph fallback: [$subject]" -Level SUCCESS
                    }
                }
            }
        }
        catch {
            Write-Log "Graph fallback could not process organized meetings: $($_.Exception.Message)" -Level WARN
            Write-Log 'Automated meeting cancellation is currently blocked in this unattended context; a delegated Exchange/Graph admin session may be required for this tenant.' -Level WARN
            $skipped++
        }
    }

    Write-Log 'Checking calendar sharing/delegate permissions on the user mailbox...'
    try {
        $calendarIdentity = "$userEmail`:\Calendar"
        $calendarPermissions = @(Get-MailboxFolderPermission -Identity $calendarIdentity -ErrorAction Stop)
        $explicitPermissions = @($calendarPermissions | Where-Object {
            $permUser = [string]$_.User
            -not [string]::IsNullOrWhiteSpace($permUser) -and
            $permUser -notin @('Default', 'Anonymous') -and
            $permUser -notmatch 'NT AUTHORITY\\SELF|Owner@Local'
        })

        foreach ($perm in $explicitPermissions) {
            $grantee = [string]$perm.User
            $rights = @($perm.AccessRights) -join ','

            if ($simulateOnly) {
                Write-Log "[DryRun] Would remove calendar permission [$rights] for [$grantee] on [$calendarIdentity]" -Level WARN
            }
            elseif ($PSCmdlet.ShouldProcess($calendarIdentity, "Remove calendar permission for [$grantee]")) {
                Remove-MailboxFolderPermission -Identity $calendarIdentity -User $grantee -Confirm:$false -ErrorAction Stop
                Write-Log "Removed calendar permission [$rights] for [$grantee] on [$calendarIdentity]" -Level SUCCESS
                $calendarPermissionsRemoved++
            }
        }
    }
    catch {
        Write-Log "Could not evaluate mailbox calendar permissions: $($_.Exception.Message)" -Level WARN
        $skipped++
    }

    Write-Log 'Checking room/resource mailboxes for booking and delegate access...'
    $resourceMailboxes = @(Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited -ErrorAction Stop)
    Write-Log "Found [$($resourceMailboxes.Count)] resource mailbox(es) to evaluate."

    foreach ($resource in $resourceMailboxes) {
        $resourceName = if ($resource.DisplayName) { $resource.DisplayName } else { [string]$resource.PrimarySmtpAddress }
        $resourceIdentity = [string]$resource.Identity
        $resourceCalendar = "$($resource.PrimarySmtpAddress)`:\Calendar"

        try {
            $calendarProcessing = Get-CalendarProcessing -Identity $resourceIdentity -ErrorAction Stop

            foreach ($propertyName in @('ResourceDelegates', 'BookInPolicy', 'RequestInPolicy', 'RequestOutOfPolicy')) {
                $split = Split-MatchedValues -Values $calendarProcessing.$propertyName -Tokens $matchTokens
                if (@($split.Removed).Count -gt 0) {
                    $removedList = @($split.Removed) -join ', '

                    if ($simulateOnly) {
                        Write-Log "[DryRun] Would remove user from [$propertyName] on resource mailbox [$resourceName]: [$removedList]" -Level WARN
                    }
                    elseif ($PSCmdlet.ShouldProcess($resourceName, "Remove [$UserPrincipalName] from [$propertyName]")) {
                        $setParams = @{
                            Identity    = $resourceIdentity
                            Confirm     = $false
                            ErrorAction = 'Stop'
                        }
                        $setParams[$propertyName] = $split.Kept
                        Set-CalendarProcessing @setParams
                        Write-Log "Removed user from [$propertyName] on resource mailbox [$resourceName]" -Level SUCCESS
                        $roomAssignmentsRemoved++
                    }
                }
            }
        }
        catch {
            Write-Log "Could not evaluate booking settings on resource mailbox [$resourceName]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }

        try {
            $roomPermissions = @(Get-MailboxFolderPermission -Identity $resourceCalendar -ErrorAction SilentlyContinue)
            $matchedPermissions = @($roomPermissions | Where-Object {
                Test-TokenMatch -Value ([string]$_.User) -Tokens $matchTokens
            })

            foreach ($perm in $matchedPermissions) {
                $grantee = [string]$perm.User
                $rights = @($perm.AccessRights) -join ','

                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove room calendar permission [$rights] for [$grantee] on [$resourceCalendar]" -Level WARN
                }
                elseif ($PSCmdlet.ShouldProcess($resourceCalendar, "Remove room calendar permission for [$grantee]")) {
                    Remove-MailboxFolderPermission -Identity $resourceCalendar -User $grantee -Confirm:$false -ErrorAction Stop
                    Write-Log "Removed room calendar permission [$rights] for [$grantee] on [$resourceCalendar]" -Level SUCCESS
                    $roomAssignmentsRemoved++
                }
            }
        }
        catch {
            Write-Log "Could not evaluate room calendar permissions on [$resourceName]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }
    }

    Write-Log "Step 06 finished. Meeting actions: [$meetingActions] | Calendar permission removals: [$calendarPermissionsRemoved] | Room access removals: [$roomAssignmentsRemoved] | Skipped/Warnings: [$skipped]"
    Write-Log 'Step 06 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 06 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
    catch {}
}
