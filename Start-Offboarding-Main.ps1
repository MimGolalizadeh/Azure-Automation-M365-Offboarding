<#
.SYNOPSIS
    Main offboarding orchestrator for Azure Automation.
.DESCRIPTION
    Runs every 12 hours in Azure Automation, reads the configured offboarding
    queue group from an Automation Variable, finds all user members in that
    group, and executes the step-by-step runbooks in order.

    The runbook waits for each child runbook to finish and also adds a
    propagation delay between steps so Exchange, Graph, licensing, and
    SharePoint changes have time to settle before the next step starts.

    After all steps succeed for a user, the user is removed from the
    offboarding queue group so they are not reprocessed on the next cycle.
.PARAMETER OffboardingGroupName
    Optional override for the Automation Variable value.
.PARAMETER ResourceGroupName
    Optional override for the Automation Variable value.
.PARAMETER AutomationAccountName
    Optional override for the Automation Variable value.
.PARAMETER SubscriptionId
    Optional Azure subscription override.
.PARAMETER DryRun
    Runs all child runbooks in DryRun mode and does not remove users from the queue group.
.PARAMETER OnlyUserPrincipalName
    Optional filter to process only one queued user, useful for testing.
.PARAMETER PollSeconds
    How often to poll child runbook jobs.
.PARAMETER MinimumStepDelaySeconds
    Minimum delay to wait after each successful step.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$OffboardingGroupName,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$AutomationAccountName,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Main-Offboarding_$(Get-Date -Format 'yyyy-MM-dd').log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [string]$OnlyUserPrincipalName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0,300)]
    [int]$PollSeconds = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0,1800)]
    [int]$MinimumStepDelaySeconds = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0,21600)]
    [int]$MaxStepWaitSeconds = 0
)

$ErrorActionPreference = 'Stop'
$script:ResolvedSubscriptionId = $null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','HEADER')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [MAIN-OFFBOARDING] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [MAIN-OFFBOARDING] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
}

function Get-ResolvedSetting {
    param(
        [string]$CurrentValue,
        [string[]]$VariableNames,
        [string]$DefaultValue = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return "$CurrentValue".Trim().Trim('"').Trim("'")
    }

    if (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue) {
        foreach ($name in $VariableNames) {
            try {
                $value = Get-AutomationVariable -Name $name -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    return "$value".Trim().Trim('"').Trim("'")
                }
            }
            catch {}
        }
    }

    return $DefaultValue
}

function Get-ResolvedIntSetting {
    param(
        [int]$CurrentValue,
        [string[]]$VariableNames,
        [int]$DefaultValue
    )

    if ($PSBoundParameters.ContainsKey('CurrentValue') -and $CurrentValue -ne 0) {
        return [int]$CurrentValue
    }

    if (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue) {
        foreach ($name in $VariableNames) {
            try {
                $value = Get-AutomationVariable -Name $name -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($value) -and ($value -as [int]) -ne $null) {
                    return [int]$value
                }
            }
            catch {}
        }
    }

    return [int]$DefaultValue
}

function Ensure-AzureContext {
    param([string]$PreferredSubscriptionId)

    if (-not (Get-Command -Name Connect-AzAccount -ErrorAction SilentlyContinue)) {
        throw 'Az.Accounts is required. Install/import Az.Accounts first.'
    }

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Log 'Connecting to Azure with managed identity...' | Out-Host
        Disable-AzContextAutosave -Scope Process -ErrorAction SilentlyContinue | Out-Null
        Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        Write-Log 'Azure connection established.' -Level SUCCESS | Out-Host
    }

    $resolvedSubscriptionId = $PreferredSubscriptionId
    if ([string]::IsNullOrWhiteSpace($resolvedSubscriptionId) -and $ctx -and $ctx.Subscription) {
        $resolvedSubscriptionId = $ctx.Subscription.Id
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedSubscriptionId)) {
        Set-AzContext -SubscriptionId $resolvedSubscriptionId -ErrorAction Stop | Out-Null
        $script:ResolvedSubscriptionId = $resolvedSubscriptionId
        Write-Log "Using Azure subscription [$resolvedSubscriptionId]" | Out-Host
    }
}

function Get-GraphAccessToken {
    Ensure-AzureContext -PreferredSubscriptionId $SubscriptionId

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
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET','DELETE')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
}

function Resolve-RunbookName {
    param(
        [string[]]$CandidateNames,
        [string[]]$AvailableRunbookNames
    )

    foreach ($candidate in $CandidateNames) {
        if ($candidate -in $AvailableRunbookNames) {
            return $candidate
        }
    }

    return $null
}

function Wait-ForAutomationJob {
    param(
        [Parameter(Mandatory = $true)][Guid]$JobId,
        [Parameter(Mandatory = $true)][string]$RunbookName
    )

    $terminalStates = @('Completed','Failed','Stopped','Suspended')
    $seenRecords = @{}
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastStatus = ''
    $lastHeartbeatBucket = -1

    do {
        $job = Get-AzAutomationJob -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Id $JobId -ErrorAction Stop

        $records = Get-AzAutomationJobOutput -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Id $JobId -Stream Any -ErrorAction SilentlyContinue
        foreach ($record in @($records)) {
            if ($record.StreamRecordId -and -not $seenRecords.ContainsKey($record.StreamRecordId)) {
                $seenRecords[$record.StreamRecordId] = $true
                try {
                    $outputRecord = Get-AzAutomationJobOutputRecord -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -JobId $JobId -Id $record.StreamRecordId -ErrorAction Stop
                    if ($outputRecord.Summary) {
                        Write-Log "[$RunbookName] $($outputRecord.Summary)" | Out-Host
                    }
                }
                catch {}
            }
        }

        $elapsedText = ('{0:mm\:ss}' -f $stopwatch.Elapsed)
        if ($job.Status -ne $lastStatus) {
            Write-Log "[$RunbookName] Current job status: [$($job.Status)] | Elapsed: [$elapsedText]" | Out-Host
            $lastStatus = [string]$job.Status
        }
        elseif ($job.Status -notin $terminalStates) {
            $heartbeatBucket = [int][Math]::Floor($stopwatch.Elapsed.TotalSeconds / 60)
            if ($heartbeatBucket -gt $lastHeartbeatBucket) {
                Write-Log "[$RunbookName] Still waiting... status: [$($job.Status)] | Elapsed: [$elapsedText]" | Out-Host
                $lastHeartbeatBucket = $heartbeatBucket
            }
        }

        if ($job.Status -in $terminalStates) {
            $stopwatch.Stop()
            return [PSCustomObject]@{
                Status  = $job.Status
                Success = ($job.Status -eq 'Completed')
            }
        }

        if ($MaxStepWaitSeconds -gt 0 -and $stopwatch.Elapsed.TotalSeconds -ge $MaxStepWaitSeconds) {
            Write-Log "[$RunbookName] Timed out after [$elapsedText] while waiting for child runbook completion." -Level ERROR | Out-Host
            $stopwatch.Stop()
            return [PSCustomObject]@{
                Status  = 'TimedOut'
                Success = $false
            }
        }

        Start-Sleep -Seconds $PollSeconds
    }
    while ($true)
}

function Get-QueuedUsers {
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $escapedName = $GroupName.Replace("'", "''")
    $groupUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escapedName'&`$select=id,displayName"
    $groupResponse = Invoke-GraphRequest -Method GET -Uri $groupUrl -AccessToken $AccessToken
    $group = @($groupResponse.value) | Select-Object -First 1

    if (-not $group) {
        throw "Offboarding group [$GroupName] was not found in Microsoft Entra ID."
    }

    Write-Log "Offboarding group resolved: [$($group.displayName)] | ID: [$($group.id)]" | Out-Host

    $users = @()
    $nextUrl = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$select=id,displayName,userPrincipalName"

    while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
        $response = Invoke-GraphRequest -Method GET -Uri $nextUrl -AccessToken $AccessToken
        foreach ($member in @($response.value)) {
            $odataType = [string]$member.'@odata.type'
            if ($odataType -eq '#microsoft.graph.user' -or -not [string]::IsNullOrWhiteSpace([string]$member.userPrincipalName)) {
                $users += [PSCustomObject]@{
                    Id                = [string]$member.id
                    DisplayName       = [string]$member.displayName
                    UserPrincipalName = [string]$member.userPrincipalName
                }
            }
        }
        $nextUrl = [string]$response.'@odata.nextLink'
    }

    return [PSCustomObject]@{
        Group = $group
        Users = @($users)
    }
}

function Remove-UserFromQueueGroup {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$UserId/`$ref"
    Invoke-GraphRequest -Method DELETE -Uri $uri -AccessToken $AccessToken | Out-Null
    Write-Log "Removed [$UserPrincipalName] from the offboarding queue group." -Level SUCCESS | Out-Host
}

$OffboardingGroupName = Get-ResolvedSetting -CurrentValue $OffboardingGroupName -VariableNames @('OffboardingGroupName','OFFBOARDINGGROUPNAME','OffboardingGroup')
$ResourceGroupName = Get-ResolvedSetting -CurrentValue $ResourceGroupName -VariableNames @('ResourceGroupName','RESOURCEGROUPNAME')
$AutomationAccountName = Get-ResolvedSetting -CurrentValue $AutomationAccountName -VariableNames @('AutomationAccountName','AUTOMATIONACCOUNTNAME')
$SubscriptionId = Get-ResolvedSetting -CurrentValue $SubscriptionId -VariableNames @('SubscriptionId','AUTOMATION_SUBSCRIPTION_ID','AzureSubscriptionId')
$MinimumStepDelaySeconds = Get-ResolvedIntSetting -CurrentValue $MinimumStepDelaySeconds -VariableNames @('OffboardingStepDelaySeconds','MainRunbookStepDelaySeconds') -DefaultValue 30
$PollSeconds = Get-ResolvedIntSetting -CurrentValue $PollSeconds -VariableNames @('OffboardingPollSeconds','MainRunbookPollSeconds') -DefaultValue 15
$MaxStepWaitSeconds = Get-ResolvedIntSetting -CurrentValue $MaxStepWaitSeconds -VariableNames @('OffboardingMaxStepWaitSeconds','MainRunbookMaxStepWaitSeconds') -DefaultValue 3600

if ([string]::IsNullOrWhiteSpace($OffboardingGroupName)) {
    throw "OffboardingGroupName is required. Create an Automation Variable named 'OffboardingGroupName'."
}
if ([string]::IsNullOrWhiteSpace($ResourceGroupName)) {
    throw "ResourceGroupName is required. Create an Automation Variable named 'ResourceGroupName'."
}
if ([string]::IsNullOrWhiteSpace($AutomationAccountName)) {
    throw "AutomationAccountName is required. Create an Automation Variable named 'AutomationAccountName'."
}

$banner = '=' * 80
Write-Log $banner -Level HEADER
Write-Log "Main offboarding run started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level HEADER
Write-Log "Automation account: [$AutomationAccountName] | Resource group: [$ResourceGroupName]" -Level HEADER
Write-Log "Offboarding group variable: [$OffboardingGroupName]" -Level HEADER
Write-Log "DryRun mode: [$([bool]$DryRun)] | PollSeconds: [$PollSeconds] | MinimumStepDelaySeconds: [$MinimumStepDelaySeconds] | MaxStepWaitSeconds: [$MaxStepWaitSeconds]" -Level HEADER
if (-not [string]::IsNullOrWhiteSpace($OnlyUserPrincipalName)) {
    Write-Log "User filter active: [$OnlyUserPrincipalName]" -Level HEADER
}
Write-Log $banner -Level HEADER

$graphToken = Get-GraphAccessToken
$publishedRunbookNames = @(Get-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -ErrorAction Stop | Select-Object -ExpandProperty Name)

$stepPlan = @(
    [PSCustomObject]@{ Id = 'STEP-01'; CandidateNames = @('Step-01_Immediate-Access-Lockdown','TEST-Step-01-Immediate-Access-Lockdown'); DelaySeconds = 30 },
    [PSCustomObject]@{ Id = 'STEP-02'; CandidateNames = @('Step-02_Convert-Mailbox-To-Shared','TEST-Step-02-Convert-Mailbox-To-Shared'); DelaySeconds = 120 },
    [PSCustomObject]@{ Id = 'STEP-03'; CandidateNames = @('Step-03_Remove-User-Groups','TEST-Step-03-Remove-User-Groups'); DelaySeconds = 180 },
    [PSCustomObject]@{ Id = 'STEP-04'; CandidateNames = @('Step-04_Remove-Distribution-And-Delegation','TEST-Step-04-Remove-Distribution-And-Delegation'); DelaySeconds = 45 },
    [PSCustomObject]@{ Id = 'STEP-05'; CandidateNames = @('Step-05_Remove-SharePoint-Access','TEST-Step-05-Remove-SharePoint-Access'); DelaySeconds = 45 },
    [PSCustomObject]@{ Id = 'STEP-06'; CandidateNames = @('Step-06_Remove-Calendar-And-Room-Access','TEST-Step-06-Remove-Calendar-And-Room-Access'); DelaySeconds = 30 },
    [PSCustomObject]@{ Id = 'STEP-07'; CandidateNames = @('Step-07_Cleanup-User-Profile','TEST-Step-07-Cleanup-User-Profile'); DelaySeconds = 30 },
    [PSCustomObject]@{ Id = 'STEP-08'; CandidateNames = @('Step-08_Remove-Remaining-Licenses','TEST-Step-08-Remove-Remaining-Licenses'); DelaySeconds = 180 },
    [PSCustomObject]@{ Id = 'STEP-09'; CandidateNames = @('Step-09_Disable-Account','TEST-Step-09-Disable-Account'); DelaySeconds = 0 }
)

foreach ($step in $stepPlan) {
    $resolvedName = Resolve-RunbookName -CandidateNames $step.CandidateNames -AvailableRunbookNames $publishedRunbookNames
    if ([string]::IsNullOrWhiteSpace($resolvedName)) {
        throw "Could not resolve an Azure Automation runbook for [$($step.Id)]. Expected one of: $($step.CandidateNames -join ', ')"
    }
    $step | Add-Member -NotePropertyName RunbookName -NotePropertyValue $resolvedName -Force
    Write-Log "[$($step.Id)] Resolved child runbook: [$resolvedName]"
}

$queueInfo = Get-QueuedUsers -GroupName $OffboardingGroupName -AccessToken $graphToken
$queueGroup = $queueInfo.Group
$queuedUsers = @($queueInfo.Users)

if (-not [string]::IsNullOrWhiteSpace($OnlyUserPrincipalName)) {
    $queuedUsers = @($queuedUsers | Where-Object { $_.UserPrincipalName -ieq $OnlyUserPrincipalName })
}

Write-Log "Detected [$($queuedUsers.Count)] user(s) currently queued for offboarding."
if ($queuedUsers.Count -eq 0) {
    Write-Log 'No queued users found. Nothing to do.' -Level SUCCESS
    exit 0
}

$successfulUsers = @()
$failedUsers = @()

foreach ($user in $queuedUsers) {
    $upn = [string]$user.UserPrincipalName
    $userId = [string]$user.Id

    Write-Log $banner -Level HEADER
    Write-Log "Processing queued user [$upn]" -Level HEADER
    Write-Log $banner -Level HEADER

    $userSucceeded = $true

    foreach ($step in $stepPlan) {
        $parameters = @{ UserPrincipalName = $upn }
        if ($DryRun) {
            $parameters['DryRun'] = $true
        }

        Write-Log "[$($step.Id)] Starting child runbook [$($step.RunbookName)] for [$upn]"
        try {
            $job = Start-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $step.RunbookName -Parameters $parameters -ErrorAction Stop
            Write-Log "[$($step.Id)] Job started: [$($job.JobId)]"
        }
        catch {
            Write-Log "[$($step.Id)] Could not start child runbook [$($step.RunbookName)]: $($_.Exception.Message)" -Level ERROR
            $userSucceeded = $false
            break
        }

        $result = Wait-ForAutomationJob -JobId $job.JobId -RunbookName $step.RunbookName
        if (-not $result.Success) {
            Write-Log "[$($step.Id)] Child runbook finished with status [$($result.Status)]. Leaving user in the queue for retry." -Level ERROR
            $userSucceeded = $false
            break
        }

        Write-Log "[$($step.Id)] Completed successfully for [$upn]" -Level SUCCESS

        $effectiveDelay = [Math]::Max([int]$MinimumStepDelaySeconds, [int]$step.DelaySeconds)
        if ($effectiveDelay -gt 0 -and $step.Id -ne 'STEP-09') {
            Write-Log "[$($step.Id)] Waiting [$effectiveDelay] second(s) before the next step to allow propagation..." -Level INFO
            Start-Sleep -Seconds $effectiveDelay
        }
    }

    if ($userSucceeded) {
        if ($DryRun) {
            Write-Log "[DryRun] Would remove [$upn] from the offboarding queue group after successful completion." -Level WARN
        }
        else {
            try {
                $freshGraphToken = Get-GraphAccessToken
                Remove-UserFromQueueGroup -GroupId $queueGroup.id -UserId $userId -UserPrincipalName $upn -AccessToken $freshGraphToken
            }
            catch {
                Write-Log "Completed all steps for [$upn], but could not remove the user from the queue group: $($_.Exception.Message)" -Level WARN
            }
        }

        $successfulUsers += $upn
    }
    else {
        $failedUsers += $upn
    }
}

Write-Log $banner -Level HEADER
Write-Log "Main offboarding run finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level HEADER
Write-Log "Succeeded user count: [$($successfulUsers.Count)]" -Level HEADER
if ($successfulUsers.Count -gt 0) {
    Write-Log "Succeeded users: $($successfulUsers -join ', ')" -Level HEADER
}
Write-Log "Failed user count: [$($failedUsers.Count)]" -Level HEADER
if ($failedUsers.Count -gt 0) {
    Write-Log "Users left in queue for retry: $($failedUsers -join ', ')" -Level HEADER
}
Write-Log $banner -Level HEADER

if ($failedUsers.Count -gt 0) {
    exit 1
}

exit 0
