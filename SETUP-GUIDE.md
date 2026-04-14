# M365 Offboarding Automation — Full Setup Guide

This guide walks through the complete installation of the Azure Automation M365 Offboarding solution
from scratch, using PowerShell. Follow each section in order.

---

## What This Solution Does

The automation reads users from a dedicated **offboarding queue group** in Entra ID and runs
them through a fixed 9-step offboarding workflow inside Azure Automation:

| Step | Runbook | Action |
|---:|---|---|
| 1 | `Step-01_Immediate-Access-Lockdown` | Revokes sign-in sessions, resets password, removes MFA/auth methods |
| 2 | `Step-02_Convert-Mailbox-To-Shared` | Converts Exchange Online mailbox to shared |
| 3 | `Step-03_Remove-User-Groups` | Removes user from all Entra ID / M365 groups |
| 4 | `Step-04_Remove-Distribution-And-Delegation` | Removes Exchange distribution and shared mailbox delegation |
| 5 | `Step-05_Remove-SharePoint-Access` | Removes direct SharePoint site access |
| 6 | `Step-06_Remove-Calendar-And-Room-Access` | Cancels meetings, removes calendar and room permissions |
| 7 | `Step-07_Cleanup-User-Profile` | Removes manager relationship and clears profile fields |
| 8 | `Step-08_Remove-Remaining-Licenses` | Removes direct M365 licenses |
| 9 | `Step-09_Disable-Account` | Disables the Entra account |

After all steps succeed the user is automatically removed from the queue group.

---

## Prerequisites

Before starting, you need:

- An Azure subscription
- A user account with at least **Contributor** access on the subscription or resource group
- **Azure PowerShell** (`Az` module) installed locally
- **Microsoft Graph PowerShell** (`Microsoft.Graph`) installed locally
- The 10 runbook `.ps1` files from this repository

### Install required local modules (if not already installed)

```powershell
Install-Module Az -Scope CurrentUser -Force -AllowClobber
Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
```

---

## Step 1 — Connect to Azure

```powershell
Connect-AzAccount -TenantId "<your-tenant-id>"
Set-AzContext -Subscription "<your-subscription-id>"
```

Verify the connection:

```powershell
Get-AzContext | Select-Object Account, Subscription, Tenant
```

---

## Step 2 — Create the Automation Account

If you already have an Automation Account, skip to Step 3.

```powershell
$rg = "IT-management-apps"        # your resource group
$aa = "m365-offboarding"          # your automation account name
$location = "westeurope"          # your preferred Azure region

# Create resource group if it doesn't exist
New-AzResourceGroup -Name $rg -Location $location -ErrorAction SilentlyContinue

# Create the Automation Account
New-AzAutomationAccount -Name $aa -ResourceGroupName $rg -Location $location -Plan Free
```

---

## Step 3 — Enable System-Assigned Managed Identity

> If you created the Automation Account in the portal with identity already enabled, verify it is on.

**Via portal:**
1. Open the Automation Account → **Identity** in the left menu
2. Under **System assigned** → set Status to **On** → **Save**
3. Copy the **Object (principal) ID** shown — you need it in Steps 6 and 7

**Via PowerShell (if not yet enabled):**

```powershell
$aa = "m365-offboarding"
$rg = "IT-management-apps"

Set-AzAutomationAccount -Name $aa -ResourceGroupName $rg -AssignSystemIdentity
```

Retrieve the Object ID:

```powershell
(Get-AzAutomationAccount -Name $aa -ResourceGroupName $rg).Identity.PrincipalId
```

Save this value — it is your `$managedIdentityId` for the next steps.

---

## Step 4 — Create the Required Automation Variables

These variables are read by the runbooks at runtime. Replace the example values with your own.

```powershell
$rg = "IT-management-apps"
$aa = "m365-offboarding"

$variables = @{
    "OffboardingGroupName"  = "AT-Offboarding-Queue"               # Entra group used as the offboarding queue
    "OffboardedGroupName"   = "AT-Offboarded-Users"                # Entra group for completed offboardings
    "AutomationAccountName" = "m365-offboarding"                   # This Automation Account's name
    "ResourceGroupName"     = "IT-management-apps"                 # Resource Group hosting this account
    "SubscriptionId"        = "00000000-0000-0000-0000-000000000000" # Your Azure subscription ID
    "ExchangeOrganization"  = "contoso.onmicrosoft.com"            # Your Exchange Online org domain
    "SharePointAdminUrl"    = "https://contoso-admin.sharepoint.com" # Your SharePoint Admin URL
}

foreach ($var in $variables.GetEnumerator()) {
    New-AzAutomationVariable -AutomationAccountName $aa -ResourceGroupName $rg `
        -Name $var.Key -Value $var.Value -Encrypted $false
    Write-Host "Created variable: $($var.Key)"
}
```

Optional variables (have sensible defaults built into the runbooks):

```powershell
# Delay in seconds between each step (default: 30)
New-AzAutomationVariable -AutomationAccountName $aa -ResourceGroupName $rg `
    -Name "OffboardingStepDelaySeconds" -Value "30" -Encrypted $false

# Polling interval in seconds while waiting for child jobs (default: 15)
New-AzAutomationVariable -AutomationAccountName $aa -ResourceGroupName $rg `
    -Name "OffboardingPollSeconds" -Value "15" -Encrypted $false

# Max wait time in seconds before timing out a child runbook (default: 3600)
New-AzAutomationVariable -AutomationAccountName $aa -ResourceGroupName $rg `
    -Name "OffboardingMaxStepWaitSeconds" -Value "3600" -Encrypted $false
```

---

## Step 5 — Import PowerShell Modules into the Automation Account

```powershell
$rg = "IT-management-apps"
$aa = "m365-offboarding"

$modules = @(
    @{ Name = "Az.Accounts";              Version = "4.0.1"  },
    @{ Name = "Az.Automation";            Version = "1.10.0" },
    @{ Name = "ExchangeOnlineManagement"; Version = "3.6.0"  },
    @{ Name = "PnP.PowerShell";           Version = "2.12.0" }
)

foreach ($mod in $modules) {
    Write-Host "Importing $($mod.Name) $($mod.Version)..."
    New-AzAutomationModule -AutomationAccountName $aa -ResourceGroupName $rg `
        -Name $mod.Name `
        -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$($mod.Name)/$($mod.Version)"
}
```

Check import status (repeat until all show `Succeeded`):

```powershell
Get-AzAutomationModule -AutomationAccountName $aa -ResourceGroupName $rg |
    Where-Object { $_.Name -in @("Az.Accounts","Az.Automation","ExchangeOnlineManagement","PnP.PowerShell") } |
    Select-Object Name, ProvisioningState, Version | Format-Table -AutoSize
```

> Module import runs in the background and typically takes 2–10 minutes per module.

---

## Step 6 — Grant Microsoft Graph Application Permissions

Connect to Microsoft Graph:

```powershell
Connect-MgGraph -TenantId "<your-tenant-id>" `
    -Scopes "AppRoleAssignment.ReadWrite.All","RoleManagement.ReadWrite.Directory","Application.ReadWrite.All","Directory.ReadWrite.All" `
    -NoWelcome
```

Grant all required Graph permissions to the managed identity:

```powershell
$managedIdentityId = "<object-id-from-step-3>"

# Get the Microsoft Graph service principal
$graphSP = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

$requiredPermissions = @(
    "User.Read.All",
    "User.ReadWrite.All",
    "User-PasswordProfile.ReadWrite.All",
    "User.RevokeSessions.All",
    "UserAuthenticationMethod.ReadWrite.All",
    "User-Phone.ReadWrite.All",
    "Group.ReadWrite.All",
    "GroupMember.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory",
    "Calendars.ReadWrite",
    "User.EnableDisableAccount.All"
)

foreach ($perm in $requiredPermissions) {
    $appRole = $graphSP.AppRoles | Where-Object {
        $_.Value -eq $perm -and $_.AllowedMemberTypes -contains "Application"
    }
    if ($appRole) {
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $managedIdentityId `
                -PrincipalId $managedIdentityId `
                -ResourceId $graphSP.Id `
                -AppRoleId $appRole.Id | Out-Null
            Write-Host "Granted: $perm"
        } catch {
            Write-Host "Already exists: $perm"
        }
    }
}
Write-Host "--- Graph permissions done ---"
```

Verify (should return 12):

```powershell
(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityId).Count
```

---

## Step 7 — Assign Entra ID Admin Roles

```powershell
$managedIdentityId = "<object-id-from-step-3>"

$roles = @{
    "User Administrator"                      = "fe930be7-5e62-47db-91af-98c3a49a38b1"
    "Exchange Administrator"                  = "29232cdf-9323-42fd-ade2-1d097af3e4de"
    "SharePoint Administrator"                = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
    "Privileged Authentication Administrator" = "7be44c8a-adaf-4e2a-84d6-ab2649e08a13"
    "Privileged Role Administrator"           = "e8611ab8-c189-46e8-94e1-60213ab1f814"
}

foreach ($roleName in $roles.Keys) {
    $templateId = $roles[$roleName]
    try {
        New-MgRoleManagementDirectoryRoleAssignment `
            -PrincipalId $managedIdentityId `
            -RoleDefinitionId $templateId `
            -DirectoryScopeId "/" | Out-Null
        Write-Host "Assigned: $roleName"
    } catch {
        Write-Host "Already assigned: $roleName"
    }
}
Write-Host "--- Entra roles done ---"
```

---

## Step 8 — Grant SharePoint Sites.FullControl.All

```powershell
$managedIdentityId = "<object-id-from-step-3>"

# Get the SharePoint Online service principal
$spoSP = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'"
$role  = $spoSP.AppRoles | Where-Object {
    $_.Value -eq "Sites.FullControl.All" -and $_.AllowedMemberTypes -contains "Application"
}

try {
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $managedIdentityId `
        -PrincipalId $managedIdentityId `
        -ResourceId $spoSP.Id `
        -AppRoleId $role.Id | Out-Null
    Write-Host "Granted: Sites.FullControl.All"
} catch {
    Write-Host "Already exists: Sites.FullControl.All"
}
```

---

## Step 9 — Grant Azure RBAC on the Automation Account

This allows the managed identity to start child runbooks and read job output.

```powershell
$rg = "IT-management-apps"
$aa = "m365-offboarding"
$managedIdentityId = "<object-id-from-step-3>"

$aaResourceId = (Get-AzAutomationAccount -Name $aa -ResourceGroupName $rg).Id

New-AzRoleAssignment `
    -ObjectId $managedIdentityId `
    -RoleDefinitionName "Contributor" `
    -Scope $aaResourceId
Write-Host "Contributor role assigned on Automation Account."
```

---

## Step 10 — Create the Entra ID Queue Groups

```powershell
# Offboarding queue — add users here to trigger offboarding
New-MgGroup `
    -DisplayName "AT-Offboarding-Queue" `
    -MailNickname "AT-Offboarding-Queue" `
    -SecurityEnabled:$true `
    -MailEnabled:$false `
    -Description "Offboarding queue — users added here are processed by the M365 offboarding automation"

# Completed offboarding archive group
New-MgGroup `
    -DisplayName "AT-Offboarded-Users" `
    -MailNickname "AT-Offboarded-Users" `
    -SecurityEnabled:$true `
    -MailEnabled:$false `
    -Description "Archive group for users who have been successfully offboarded"

Write-Host "Groups created."
```

> If these groups already exist, this command will error — that is expected. Verify they exist with:
> `Get-MgGroup -Filter "DisplayName eq 'AT-Offboarding-Queue'" | Select-Object Id, DisplayName`

---

## Step 11 — Upload and Publish All Runbooks

```powershell
$rg     = "IT-management-apps"
$aa     = "m365-offboarding"
$folder = "C:\path\to\your\runbook\files"   # folder containing the .ps1 files

$runbooks = @(
    "Start-Offboarding-Main",
    "Step-01_Immediate-Access-Lockdown",
    "Step-02_Convert-Mailbox-To-Shared",
    "Step-03_Remove-User-Groups",
    "Step-04_Remove-Distribution-And-Delegation",
    "Step-05_Remove-SharePoint-Access",
    "Step-06_Remove-Calendar-And-Room-Access",
    "Step-07_Cleanup-User-Profile",
    "Step-08_Remove-Remaining-Licenses",
    "Step-09_Disable-Account"
)

foreach ($rb in $runbooks) {
    $path = Join-Path $folder "$rb.ps1"
    Import-AzAutomationRunbook -AutomationAccountName $aa -ResourceGroupName $rg `
        -Path $path -Name $rb -Type PowerShell -Force
    Publish-AzAutomationRunbook -AutomationAccountName $aa -ResourceGroupName $rg -Name $rb
    Write-Host "Published: $rb"
}
```

Verify all are published:

```powershell
Get-AzAutomationRunbook -AutomationAccountName $aa -ResourceGroupName $rg |
    Select-Object Name, State | Sort-Object Name | Format-Table -AutoSize
```

All runbooks should show `Published`.

---

## Step 12 — (Optional) Create the PowerShell 7.4 Runtime for Step-05

Step-05 (SharePoint) can optionally use a dedicated PowerShell 7.4 runtime with PnP.PowerShell. The default tested deployment remains classic Azure Automation `PowerShell` for all runbooks.

```powershell
$sub = "<your-subscription-id>"
$rg  = "IT-management-apps"
$aa  = "m365-offboarding"
$rt  = "PowerShell-74-Custom"
$api = "2023-05-15-preview"

# Get a fresh token (Az module returns SecureString — convert it)
$secureToken = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
$token = [System.Net.NetworkCredential]::new("", $secureToken).Password
$h    = @{ "Authorization" = "Bearer $token"; "Content-Type" = "application/json" }
$base = "https://management.azure.com/subscriptions/$sub/resourceGroups/$rg/providers/Microsoft.Automation/automationAccounts/$aa"

# Create the runtime environment
$r1 = Invoke-RestMethod -Uri "$base/runtimeEnvironments/$($rt)?api-version=$api" `
    -Method PUT -Headers $h `
    -Body '{"properties":{"runtime":{"language":"PowerShell","version":"7.4"},"description":"PS 7.4 runtime for PnP.PowerShell"}}'
Write-Host "Runtime created: $($r1.name)"

# Add PnP.PowerShell as a package in the runtime
Start-Sleep -Seconds 5
$pkgUri  = "$base/runtimeEnvironments/$rt/packages/PnP.PowerShell?api-version=$api"
$pkgBody = '{"properties":{"contentLink":{"uri":"https://www.powershellgallery.com/api/v2/package/PnP.PowerShell/"}}}'
$pkg = Invoke-RestMethod -Uri $pkgUri -Method PUT -Headers $h -Body $pkgBody
Write-Host "Package added: $($pkg.name) — State: $($pkg.properties.provisioningState)"
```

Check when the package is ready (repeat until `Succeeded`):

```powershell
$status = Invoke-RestMethod -Uri $pkgUri -Method GET -Headers $h
Write-Host "PnP.PowerShell: $($status.properties.provisioningState) — Version: $($status.properties.version)"
```

Link Step-05 to this runtime:

```powershell
$rbUri = "$base/runbooks/Step-05_Remove-SharePoint-Access?api-version=$api"
$rb    = Invoke-RestMethod -Uri $rbUri -Method GET -Headers $h
$rb.properties | Add-Member -MemberType NoteProperty -Name "runtimeEnvironment" -Value $rt -Force
$result = Invoke-RestMethod -Uri $rbUri -Method PATCH -Headers $h -Body ($rb | ConvertTo-Json -Depth 10)
Write-Host "Step-05 linked to: $($result.properties.runtimeEnvironment)"
```

---

## Step 13 — Create the 12-Hour Schedule

```powershell
$rg = "IT-management-apps"
$aa = "m365-offboarding"

# Start time: next upcoming 00:00 or 12:00
$startTime = (Get-Date).Date.AddHours(([Math]::Ceiling((Get-Date).Hour / 12) * 12))
if ($startTime -le (Get-Date)) { $startTime = $startTime.AddHours(12) }

New-AzAutomationSchedule -AutomationAccountName $aa -ResourceGroupName $rg `
    -Name "Offboarding-Every-12-Hours" `
    -StartTime $startTime `
    -HourInterval 12 `
    -Description "Runs the M365 offboarding orchestrator every 12 hours"

# Link the schedule to the main runbook
Register-AzAutomationScheduledRunbook -AutomationAccountName $aa -ResourceGroupName $rg `
    -RunbookName "Start-Offboarding-Main" `
    -ScheduleName "Offboarding-Every-12-Hours"

Write-Host "Schedule created and linked."
```

Verify:

```powershell
Get-AzAutomationSchedule -AutomationAccountName $aa -ResourceGroupName $rg |
    Select-Object Name, IsEnabled, Interval, Frequency, NextRun | Format-Table -AutoSize
```

---

## Step 14 — Test With a Dry Run

Before going live, run a dry-run test on a single user. This makes **no changes** — it only logs what would happen.

```powershell
$rg = "IT-management-apps"
$aa = "m365-offboarding"

$params = @{
    DryRun                = $true
    OnlyUserPrincipalName = "testuser@contoso.com"   # replace with a real test UPN
}

Start-AzAutomationRunbook -AutomationAccountName $aa -ResourceGroupName $rg `
    -Name "Start-Offboarding-Main" `
    -Parameters $params
```

Monitor the job output in the portal:
**Automation Account → Jobs → select the job → Output / All Logs**

---

## Day-to-Day Usage

### Queue a user for offboarding

Add the user to the **`AT-Offboarding-Queue`** group in Entra ID.
The automation picks them up on the next scheduled run (every 12 hours).

**Via portal:** Entra ID → Groups → AT-Offboarding-Queue → Members → + Add members

**Via PowerShell:**

```powershell
Connect-MgGraph -Scopes "GroupMember.ReadWrite.All" -NoWelcome

$groupId = (Get-MgGroup -Filter "DisplayName eq 'AT-Offboarding-Queue'").Id
$userId  = (Get-MgUser -Filter "UserPrincipalName eq 'user@contoso.com'").Id

New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId
Write-Host "User queued for offboarding."
```

### Run immediately (without waiting for the schedule)

```powershell
Start-AzAutomationRunbook -AutomationAccountName "m365-offboarding" `
    -ResourceGroupName "IT-management-apps" `
    -Name "Start-Offboarding-Main"
```

### Run a single step manually for testing

```powershell
$params = @{ UserPrincipalName = "user@contoso.com"; dryRun = $true }

Start-AzAutomationRunbook -AutomationAccountName "m365-offboarding" `
    -ResourceGroupName "IT-management-apps" `
    -Name "Step-01_Immediate-Access-Lockdown" `
    -Parameters $params
```

---

## Quick Setup Checklist

- [ ] **Step 1** — Connect to Azure (`Connect-AzAccount`)
- [ ] **Step 2** — Create Automation Account (or confirm it exists)
- [ ] **Step 3** — Enable system-assigned managed identity, note the Object ID
- [ ] **Step 4** — Create all Automation Variables
- [ ] **Step 5** — Import modules: `Az.Accounts`, `Az.Automation`, `ExchangeOnlineManagement`, `PnP.PowerShell`
- [ ] **Step 6** — Grant 12 Microsoft Graph application permissions
- [ ] **Step 7** — Assign 5 Entra admin roles
- [ ] **Step 8** — Grant SharePoint `Sites.FullControl.All`
- [ ] **Step 9** — Assign `Contributor` RBAC on the Automation Account
- [ ] **Step 10** — Create `AT-Offboarding-Queue` and `AT-Offboarded-Users` groups in Entra ID
- [ ] **Step 11** — Upload and publish all 10 runbooks
- [ ] **Step 12** — Create `PowerShell-74-Custom` runtime, add `PnP.PowerShell`, link to Step-05
- [ ] **Step 13** — Create `Offboarding-Every-12-Hours` schedule linked to `Start-Offboarding-Main`
- [ ] **Step 14** — Run a dry-run test before going live

---

## Automation Variables Reference

| Variable | Required | Example Value | Purpose |
|---|---|---|---|
| `OffboardingGroupName` | Yes | `AT-Offboarding-Queue` | Queue group — users here get offboarded |
| `OffboardedGroupName` | Yes | `AT-Offboarded-Users` | Archive group for completed offboardings |
| `AutomationAccountName` | Yes | `m365-offboarding` | Automation Account name |
| `ResourceGroupName` | Yes | `IT-management-apps` | Resource Group name |
| `SubscriptionId` | Yes | `00000000-...` | Azure subscription ID |
| `ExchangeOrganization` | Yes | `contoso.onmicrosoft.com` | Exchange Online org domain |
| `SharePointAdminUrl` | Yes | `https://contoso-admin.sharepoint.com` | SharePoint Admin URL |
| `OffboardingStepDelaySeconds` | Optional | `30` | Delay between steps (seconds) |
| `OffboardingPollSeconds` | Optional | `15` | Child job polling interval (seconds) |
| `OffboardingMaxStepWaitSeconds` | Optional | `3600` | Max wait time per child runbook (seconds) |

---

## Graph Permissions Reference

| Permission | Used By | Why |
|---|---|---|
| `User.Read.All` | Main, 01, 07, 08, 09 | Look up users and read account state |
| `User.ReadWrite.All` | 07, 09 | Update profile fields and disable account |
| `User-PasswordProfile.ReadWrite.All` | 01 | Reset password |
| `User.RevokeSessions.All` | 01 | Revoke sign-in sessions |
| `UserAuthenticationMethod.ReadWrite.All` | 01 | Remove MFA / Authenticator / FIDO methods |
| `User-Phone.ReadWrite.All` | 07 | Clear phone fields |
| `Group.ReadWrite.All` | Main, 03 | Read and update groups |
| `GroupMember.ReadWrite.All` | Main, 03 | Remove user from groups |
| `Directory.ReadWrite.All` | 03, 07 | Broad directory operations |
| `RoleManagement.ReadWrite.Directory` | 03 | Remove from role-assignable groups |
| `Calendars.ReadWrite` | 06 | Calendar cleanup |
| `User.EnableDisableAccount.All` | 09 | Disable account |

---

## Entra Admin Roles Reference

| Role | Template ID | Why |
|---|---|---|
| `User Administrator` | `fe930be7-5e62-47db-91af-98c3a49a38b1` | Update user properties, disable account |
| `Exchange Administrator` | `29232cdf-9323-42fd-ade2-1d097af3e4de` | Connect to Exchange Online |
| `SharePoint Administrator` | `f28a1f50-f6e7-4571-818b-6a12f2af6b6c` | Remove SharePoint access |
| `Privileged Authentication Administrator` | `7be44c8a-adaf-4e2a-84d6-ab2649e08a13` | Clear protected auth data |
| `Privileged Role Administrator` | `e8611ab8-c189-46e8-94e1-60213ab1f814` | Remove role-assignable group memberships |

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Step-05 fails with `PnP.PowerShell` not found | Runbook not linked to PS 7.4 runtime | Repeat Step 12 |
| `Insufficient privileges` in Step-01/03/09 | Graph permission missing or not consented | Re-run Step 6 |
| Exchange steps fail with auth error | Managed identity not permitted for Exchange | Verify `Exchange Administrator` role is assigned |
| Main runbook can't start child jobs | RBAC not assigned on Automation Account | Re-run Step 9 |
| Variables not found at runtime | Variable name mismatch | Check exact variable names in Step 4 |
| User not picked up by scheduler | User not in the queue group | Add user to `AT-Offboarding-Queue` in Entra ID |
