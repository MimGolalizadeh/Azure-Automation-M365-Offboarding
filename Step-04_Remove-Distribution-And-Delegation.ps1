<#
.SYNOPSIS
    Step 04 | Remove Distribution Memberships and Shared Mailbox Delegation
.DESCRIPTION
    Removes the user from:
      - distribution lists
      - mail-enabled security groups
      - shared mailbox permissions/delegation where the user has access

    This step is intentionally limited to Exchange-side membership and mailbox
    delegation cleanup.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER ExchangeOrganization
    Exchange Online organization value, for example: contoso.onmicrosoft.com
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
    [string]$LogPath = ".\Logs\Step-04_$(Get-Date -Format 'yyyy-MM-dd').log",

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
    $entry = "[$timestamp] [$Level] [STEP-04] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-04] Could not write log file [$LogPath]: $($_.Exception.Message)"
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

function Get-SafeArray {
    param([object]$Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

try {
    Write-Log "Starting Exchange cleanup step for [$UserPrincipalName]"

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

    $recipient = $null
    foreach ($identity in @($UserPrincipalName, $UserPrincipalName.ToLower())) {
        try {
            $recipient = Get-Recipient -Identity $identity -ErrorAction Stop
            break
        }
        catch {}
    }

    if (-not $recipient) {
        try {
            $mailboxCandidate = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            $recipient = Get-Recipient -Identity $mailboxCandidate.PrimarySmtpAddress.ToString() -ErrorAction Stop
        }
        catch {}
    }

    if (-not $recipient) {
        Write-Log "No Exchange recipient could be resolved for [$UserPrincipalName]. Falling back to UPN-based matching for Exchange cleanup." -Level WARN
        $userEmail = [string]$UserPrincipalName
        $userDn = ''
        $userAlias = ($UserPrincipalName -split '@')[0]
        $userId = ''
        $matchTokens = @($UserPrincipalName, $UserPrincipalName.ToLower(), $userEmail, $userAlias) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    }
    else {
        $userEmail = [string]$recipient.PrimarySmtpAddress
        $userDn = [string]$recipient.DistinguishedName
        $userAlias = [string]$recipient.Alias
        $userId = [string]$recipient.ExternalDirectoryObjectId
        $matchTokens = @($UserPrincipalName, $UserPrincipalName.ToLower(), $userEmail, $userDn, $userAlias, $userId) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        Write-Log "Recipient found: [$($recipient.DisplayName)] | Email: [$userEmail] | Type: [$($recipient.RecipientType)]"
    }

    $removedMember = 0
    $removedDelegation = 0
    $skipped = 0

    Write-Log 'Checking distribution groups...'
    $distributionGroups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Where-Object { $_.RecipientTypeDetails -ne 'MailUniversalSecurityGroup' })
    foreach ($group in $distributionGroups) {
        try {
            $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $memberEmails = @($members | ForEach-Object { [string]$_.PrimarySmtpAddress })
            if ($memberEmails -contains $userEmail) {
                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove user from distribution group: [$($group.DisplayName)]" -Level WARN
                }
                elseif ($PSCmdlet.ShouldProcess($group.DisplayName, 'Remove user from distribution group')) {
                    Remove-DistributionGroupMember -Identity $group.Identity -Member $UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                    Write-Log "Removed user from distribution group: [$($group.DisplayName)]" -Level SUCCESS
                    $removedMember++
                }
            }
        }
        catch {
            Write-Log "Could not process distribution group [$($group.DisplayName)]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }
    }

    Write-Log 'Checking mail-enabled security groups...'
    $mailSecurityGroups = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop)
    foreach ($group in $mailSecurityGroups) {
        try {
            $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $memberEmails = @($members | ForEach-Object { [string]$_.PrimarySmtpAddress })
            if ($memberEmails -contains $userEmail) {
                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove user from mail-enabled security group: [$($group.DisplayName)]" -Level WARN
                }
                elseif ($PSCmdlet.ShouldProcess($group.DisplayName, 'Remove user from mail-enabled security group')) {
                    Remove-DistributionGroupMember -Identity $group.Identity -Member $UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                    Write-Log "Removed user from mail-enabled security group: [$($group.DisplayName)]" -Level SUCCESS
                    $removedMember++
                }
            }
        }
        catch {
            Write-Log "Could not process mail-enabled security group [$($group.DisplayName)]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }
    }

    Write-Log 'Checking shared mailbox delegation...'
    $sharedMailboxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
    foreach ($mailbox in $sharedMailboxes) {
        $mailboxName = if ($mailbox.DisplayName) { $mailbox.DisplayName } else { $mailbox.Alias }

        try {
            $permissions = @(Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue)
            $fullAccessPermission = $permissions | Where-Object {
                $permUser = [string]$_.User
                (-not $_.IsInherited) -and ($_.AccessRights -contains 'FullAccess') -and ($matchTokens | Where-Object { $permUser -like "*$_*" })
            }

            if ($fullAccessPermission) {
                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove FullAccess delegation from shared mailbox: [$mailboxName]" -Level WARN
                }
                elseif ($PSCmdlet.ShouldProcess($mailboxName, 'Remove shared mailbox FullAccess delegation')) {
                    Remove-MailboxPermission -Identity $mailbox.Identity -User $UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                    Write-Log "Removed FullAccess delegation from shared mailbox: [$mailboxName]" -Level SUCCESS
                    $removedDelegation++
                }
            }
        }
        catch {
            Write-Log "Could not process FullAccess delegation on shared mailbox [$mailboxName]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }

        try {
            $recipientPerms = @(Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue)
            $sendAsPermission = $recipientPerms | Where-Object {
                $trustee = [string]$_.Trustee
                ($_.AccessRights -contains 'SendAs') -and ($matchTokens | Where-Object { $trustee -like "*$_*" })
            }

            if ($sendAsPermission) {
                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove SendAs delegation from shared mailbox: [$mailboxName]" -Level WARN
                }
                elseif ($PSCmdlet.ShouldProcess($mailboxName, 'Remove shared mailbox SendAs delegation')) {
                    Remove-RecipientPermission -Identity $mailbox.Identity -Trustee $UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                    Write-Log "Removed SendAs delegation from shared mailbox: [$mailboxName]" -Level SUCCESS
                    $removedDelegation++
                }
            }
        }
        catch {
            Write-Log "Could not process SendAs delegation on shared mailbox [$mailboxName]: $($_.Exception.Message)" -Level WARN
            $skipped++
        }
    }

    Write-Log "Step 04 finished. Membership removals: [$removedMember] | Delegation removals: [$removedDelegation] | Skipped/Warnings: [$skipped]"
    Write-Log 'Step 04 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 04 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
    catch {}
}
