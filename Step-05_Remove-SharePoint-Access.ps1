<#
.SYNOPSIS
    Step 05 | Remove SharePoint Access
.DESCRIPTION
    Removes the user from SharePoint Online site collections where they have
    direct access, group membership, or site collection admin access.

    This step is intentionally focused on SharePoint site access cleanup only.
.PARAMETER UserPrincipalName
    UPN of the user to process.
.PARAMETER SharePointAdminUrl
    SharePoint admin URL, for example: https://contoso-admin.sharepoint.com
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
    [string]$SharePointAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Logs\Step-05_$(Get-Date -Format 'yyyy-MM-dd').log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl) -and (Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue)) {
    foreach ($name in @('SharePointAdminUrl', 'SHAREPOINTADMINURL', 'SharepointAdminUrl')) {
        try {
            $val = Get-AutomationVariable -Name $name -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $SharePointAdminUrl = "$val"
                break
            }
        }
        catch {}
    }
}

$SharePointAdminUrl = "$SharePointAdminUrl".Trim().Trim('"').Trim("'").TrimEnd('/')
if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
    throw "SharePointAdminUrl is required. Pass -SharePointAdminUrl or create an Azure Automation variable named 'SharePointAdminUrl'."
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [STEP-05] $Message"
    Write-Output $entry

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
        }
        Add-Content -Path $LogPath -Value $entry -WhatIf:$false
    }
    catch {
        Write-Output "[$timestamp] [WARN] [STEP-05] Could not write log file [$LogPath]: $($_.Exception.Message)"
    }
}

function Import-PnPPowerShellModule {
    $candidateModules = @(
        'PnP.PowerShell',
        'C:\usr\src\PSModules\PnP.PowerShell\PnP.PowerShell.psd1',
        'C:\Modules\User\PnP.PowerShell\PnP.PowerShell.psd1',
        'C:\Modules\Global\PnP.PowerShell\PnP.PowerShell.psd1'
    )

    foreach ($candidate in $candidateModules) {
        try {
            Import-Module $candidate -ErrorAction Stop
            Write-Log "Loaded PnP.PowerShell from [$candidate]." -Level SUCCESS
            return
        }
        catch {}
    }

    throw 'PnP.PowerShell could not be loaded. Ensure the module is available in the Azure Automation runtime or custom runtime environment.'
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

    return [string]$token
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    Invoke-RestMethod -Method GET -Uri $Uri -Headers @{ Authorization = "Bearer $AccessToken" } -ErrorAction Stop
}

function Test-SPOUserMatch {
    param(
        [Parameter(Mandatory = $true)]$SpoUser,
        [Parameter(Mandatory = $true)][string]$TargetUpn,
        [Parameter(Mandatory = $false)][string]$TargetObjectId
    )

    $candidates = @()
    if ($TargetUpn) {
        $normalizedUpn = $TargetUpn.ToLower()
        $candidates += $normalizedUpn
        $candidates += "i:0#.f|membership|$normalizedUpn"
        $candidates += "i:0#.w|$normalizedUpn"
    }
    if ($TargetObjectId) {
        $normalizedObjectId = $TargetObjectId.ToLower()
        $candidates += $normalizedObjectId
        $candidates += "c:0o.c|federateddirectoryclaimprovider|$normalizedObjectId"
    }
    $candidates = @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    $sourceValues = @(
        [string]$SpoUser.LoginName,
        [string]$SpoUser.Email,
        [string]$SpoUser.UserPrincipalName
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLower() }

    foreach ($value in $sourceValues) {
        if ($candidates -contains $value) { return $true }
        foreach ($candidate in $candidates) {
            if ($value -like "*$candidate*") { return $true }
        }
    }

    return $false
}

try {
    Write-Log "Starting SharePoint cleanup for [$UserPrincipalName]"

    $simulateOnly = [bool]$DryRun -or [bool]$WhatIfPreference
    if ($simulateOnly) {
        Write-Log 'Dry-run mode is enabled. No SharePoint changes will be made.' -Level WARN
    }

    $env:PNPPOWERSHELL_UPDATECHECK = 'Off'
    Write-Log "Running in PowerShell version [$($PSVersionTable.PSVersion.ToString())]"
    Import-PnPPowerShellModule

    Write-Log 'Connecting to SharePoint admin via managed identity...'
    $adminConnection = Connect-PnPOnline -Url $SharePointAdminUrl -ManagedIdentity -ReturnConnection -ErrorAction Stop
    Write-Log 'Connected to SharePoint admin.' -Level SUCCESS

    $userObjectId = $null
    try {
        $graphToken = Get-GraphAccessToken
        $graphUser = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=id,userPrincipalName" -AccessToken $graphToken
        $userObjectId = $graphUser.id
        if ($userObjectId) {
            Write-Log 'User object ID resolved from Microsoft Graph.' -Level SUCCESS
        }
    }
    catch {
        Write-Log "Could not resolve Graph object ID. Matching will continue using UPN patterns only: $($_.Exception.Message)" -Level WARN
    }

    Write-Log 'Enumerating SharePoint site collections...'
    $sites = @(Get-PnPTenantSite -Connection $adminConnection -ErrorAction Stop | Where-Object {
        $_.Url -notmatch '-my\.sharepoint\.com/personal/'
    })
    Write-Log "Found [$($sites.Count)] site collection(s) to evaluate."

    $removedCount = 0
    $skippedCount = 0

    foreach ($site in $sites) {
        try {
            $siteConnection = Connect-PnPOnline -Url $site.Url -ManagedIdentity -ReturnConnection -ErrorAction Stop
            $siteUsers = @(Get-PnPUser -Connection $siteConnection -ErrorAction Stop)
            $matchedUsers = @($siteUsers | Where-Object {
                Test-SPOUserMatch -SpoUser $_ -TargetUpn $UserPrincipalName -TargetObjectId $userObjectId
            })

            foreach ($matchedUser in $matchedUsers) {
                $loginName = [string]$matchedUser.LoginName
                if ([string]::IsNullOrWhiteSpace($loginName)) { continue }

                if ($simulateOnly) {
                    Write-Log "[DryRun] Would remove SharePoint user [$loginName] from [$($site.Url)]" -Level WARN
                    continue
                }

                if ($PSCmdlet.ShouldProcess($site.Url, "Remove SharePoint user [$loginName]")) {
                    if ($matchedUser.IsSiteAdmin) {
                        Remove-PnPSiteCollectionAdmin -Owners $loginName -Connection $siteConnection -ErrorAction SilentlyContinue
                    }
                    Remove-PnPUser -Identity $loginName -Connection $siteConnection -Force -ErrorAction Stop
                    Write-Log "Removed SharePoint user [$loginName] from [$($site.Url)]" -Level SUCCESS
                    $removedCount++
                }
            }
        }
        catch {
            Write-Log "Could not process site [$($site.Url)]: $($_.Exception.Message)" -Level WARN
            $skippedCount++
        }
    }

    Write-Log "Step 05 finished. SharePoint removals: [$removedCount] | Skipped/Warnings: [$skippedCount]"
    Write-Log 'Step 05 completed successfully.' -Level SUCCESS
}
catch {
    Write-Log "Step 05 failed: $($_.Exception.Message)" -Level ERROR
    throw
}
finally {
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    }
    catch {}
}
