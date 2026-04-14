# Azure Automation M365 Offboarding

A modular PowerShell solution for **Azure Automation** that offboards Microsoft 365 users in a controlled, auditable 9-step workflow.

This repository is designed for organizations that want to automate user offboarding across **Microsoft Entra ID**, **Exchange Online**, **SharePoint Online**, and **Microsoft 365 licensing**.

---

## What this solution does

The automation:

- reads users from a dedicated **offboarding queue group**
- runs each offboarding task in a fixed sequence
- waits between steps so Microsoft 365 changes can propagate cleanly
- supports **dry-run testing** before making changes
- writes detailed log output for each runbook
- removes the user from the queue group only after all steps succeed

The main orchestrator is:

- `Start-Offboarding-Main.ps1`

The child runbooks are:

- `Step-01_Immediate-Access-Lockdown.ps1`
- `Step-02_Convert-Mailbox-To-Shared.ps1`
- `Step-03_Remove-User-Groups.ps1`
- `Step-04_Remove-Distribution-And-Delegation.ps1`
- `Step-05_Remove-SharePoint-Access.ps1`
- `Step-06_Remove-Calendar-And-Room-Access.ps1`
- `Step-07_Cleanup-User-Profile.ps1`
- `Step-08_Remove-Remaining-Licenses.ps1`
- `Step-09_Disable-Account.ps1`

---

## Workflow overview

| Step | Runbook | Purpose |
|---:|---|---|
| 1 | `Step-01_Immediate-Access-Lockdown` | Revokes sign-in sessions, resets the password, and removes registered authentication methods. |
| 2 | `Step-02_Convert-Mailbox-To-Shared` | Converts the user's Exchange Online mailbox to a shared mailbox. |
| 3 | `Step-03_Remove-User-Groups` | Removes the user from Entra ID / Microsoft 365 groups except the protected offboarding queue group. |
| 4 | `Step-04_Remove-Distribution-And-Delegation` | Removes Exchange distribution memberships, mail-enabled security group memberships, and shared mailbox delegation. |
| 5 | `Step-05_Remove-SharePoint-Access` | Removes direct SharePoint site access and site collection admin access. |
| 6 | `Step-06_Remove-Calendar-And-Room-Access` | Cancels future meetings, removes calendar permissions, and clears room/resource access. |
| 7 | `Step-07_Cleanup-User-Profile` | Removes the manager relationship and clears non-essential profile fields. |
| 8 | `Step-08_Remove-Remaining-Licenses` | Removes remaining direct Microsoft 365 licenses and reports group-based assignments still propagating. |
| 9 | `Step-09_Disable-Account` | Disables the Microsoft Entra account after the earlier steps complete. |

---

## Requirements and access

### Core Azure Automation requirements

- An **Azure Automation Account** with **system-assigned managed identity enabled**
- Published PowerShell runbooks for the main script and all 9 steps
- A schedule for `Start-Offboarding-Main` if you want unattended runs
- Azure Automation variables configured for your tenant

### Required Automation variables

| Variable | Purpose |
|---|---|
| `OffboardingGroupName` | The Entra ID group used as the offboarding queue |
| `ResourceGroupName` | The resource group that contains the Automation Account |
| `AutomationAccountName` | The Automation Account name |
| `SubscriptionId` | Azure subscription used by the runbooks |
| `ExchangeOrganization` | Exchange Online organization, e.g. `contoso.onmicrosoft.com` |
| `SharePointAdminUrl` | SharePoint admin URL, e.g. `https://contoso-admin.sharepoint.com` |

### Required PowerShell modules

| Area | Modules |
|---|---|
| Main orchestration | `Az.Accounts`, `Az.Automation` |
| Graph-based steps | `Az.Accounts` |
| Exchange steps | `Az.Accounts`, `ExchangeOnlineManagement` |
| SharePoint step | `Az.Accounts`, `PnP.PowerShell` |

> Important: this solution is tested with classic Azure Automation `PowerShell` runbooks for all 10 scripts. If you choose `PowerShell 7.x` or a custom runtime, you must explicitly provision that runtime and install the required modules there.

> Also ensure the `SubscriptionId` Automation Variable is set to the subscription containing the Automation Account.

### Required access and roles

At a minimum, the Automation Account managed identity should have:

- **Microsoft Graph application permissions** such as:
  - `User.Read.All`
  - `User.ReadWrite.All`
  - `User-PasswordProfile.ReadWrite.All`
  - `User.RevokeSessions.All`
  - `UserAuthenticationMethod.ReadWrite.All`
  - `User-Phone.ReadWrite.All`
  - `Group.ReadWrite.All`
  - `GroupMember.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `Calendars.ReadWrite`
  - `RoleManagement.ReadWrite.Directory` *(required when role-assignable groups must be removed)*
  - `User.EnableDisableAccount.All` *(recommended / required in some tenants for Step 09 account disablement)*
- **SharePoint application permission**:
  - `Sites.FullControl.All`
- **Exchange Online app-only / managed identity access**:
  - the managed identity must be allowed to connect to Exchange Online for `Step-02`, `Step-04`, and `Step-06`
  - in some tenants this is represented as `Exchange.ManageAsApp` or equivalent Exchange app-only access configuration
- **Entra / Microsoft 365 admin roles** such as:
  - `User Administrator`
  - `Exchange Administrator`
  - `SharePoint Administrator`
  - `Privileged Authentication Administrator`
  - `Privileged Role Administrator` *(if role-assignable groups are involved)*
- **Azure RBAC** on the Automation Account or hosting resource group:
  - `Automation Contributor` or `Contributor`
  - this is needed so `Start-Offboarding-Main` can start child runbooks and read job output/status

For the full permission matrix, runtime notes, and setup checklist, see:

- [`AZURE-AUTOMATION-SETUP.md`](./AZURE-AUTOMATION-SETUP.md)

---

## Step-by-step installation

### 1) Create the Azure Automation Account

1. Create or choose an **Azure Automation Account**.
2. Enable the **system-assigned managed identity**.
3. Confirm the account is using the correct Azure subscription and tenant.

### 2) Import the required modules

Import these into the Automation Account:

- `Az.Accounts`
- `Az.Automation`
- `ExchangeOnlineManagement`
- `PnP.PowerShell`

### 3) Configure the runtime for SharePoint

For `Step-05_Remove-SharePoint-Access`:

1. Create or use a **PowerShell 7.4** runtime environment.
2. Ensure `PnP.PowerShell` is available in that runtime.
3. Link the Step 05 runbook to that runtime.

### 4) Grant Microsoft Graph permissions and admin consent

Grant the Automation Account managed identity the Graph application permissions required by the steps, including user, group, authentication method, calendar, and license management access.

Be sure to include, where applicable:

- `User.Read.All`
- `User.ReadWrite.All`
- `User-PasswordProfile.ReadWrite.All`
- `User.RevokeSessions.All`
- `UserAuthenticationMethod.ReadWrite.All`
- `User-Phone.ReadWrite.All`
- `Group.ReadWrite.All`
- `GroupMember.ReadWrite.All`
- `Directory.ReadWrite.All`
- `Calendars.ReadWrite`
- `RoleManagement.ReadWrite.Directory` *(if users may belong to role-assignable groups)*
- `User.EnableDisableAccount.All` *(recommended / sometimes required for account disablement)*

Then grant **admin consent**.

### 5) Assign Entra and Microsoft 365 admin roles

Assign the managed identity the roles needed by the runbooks:

- `User Administrator`
- `Exchange Administrator`
- `SharePoint Administrator`
- `Privileged Authentication Administrator`
- `Privileged Role Administrator` *(only if role-assignable groups are in scope)*

### 6) Grant service-specific and Azure RBAC access

1. Grant SharePoint application access: `Sites.FullControl.All`.
2. Grant Azure RBAC on the Automation Account or its resource group:
   - `Automation Contributor` or `Contributor`
3. Verify that the main runbook can call:
   - `Start-AzAutomationRunbook`
   - `Get-AzAutomationJob`
   - `Get-AzAutomationJobOutput`

### 7) Create all required Automation variables

Create these variables in the Automation Account:

| Variable | Required | Purpose |
|---|---|---|
| `OffboardingGroupName` | Yes | Queue group containing users to process |
| `ResourceGroupName` | Yes | Resource group hosting the Automation Account |
| `AutomationAccountName` | Yes | Automation Account name |
| `SubscriptionId` | Yes | Azure subscription context |
| `ExchangeOrganization` | Yes | Exchange Online organization value |
| `SharePointAdminUrl` | Yes | SharePoint admin endpoint |
| `OffboardingStepDelaySeconds` | Optional | Delay between steps |
| `OffboardingPollSeconds` | Optional | Child-job polling interval |
| `OffboardingMaxStepWaitSeconds` | Optional | Timeout while waiting for a child runbook |

### 8) Upload and publish the runbooks

Import all `.ps1` files into Azure Automation as **PowerShell runbooks**:

- `Start-Offboarding-Main.ps1`
- `Step-01_Immediate-Access-Lockdown.ps1`
- `Step-02_Convert-Mailbox-To-Shared.ps1`
- `Step-03_Remove-User-Groups.ps1`
- `Step-04_Remove-Distribution-And-Delegation.ps1`
- `Step-05_Remove-SharePoint-Access.ps1`
- `Step-06_Remove-Calendar-And-Room-Access.ps1`
- `Step-07_Cleanup-User-Profile.ps1`
- `Step-08_Remove-Remaining-Licenses.ps1`
- `Step-09_Disable-Account.ps1`

After import, **publish** each runbook.

### 9) Create the offboarding queue group

1. Create an Entra ID security group, for example `Offboarding-Queue`.
2. Set the same name in the `OffboardingGroupName` Automation variable.
3. Add users to this group when they need to be offboarded.

### 10) Create the schedule

Set `Start-Offboarding-Main` to run on the schedule you want.

A common pattern is **every 12 hours**.

### 11) Test in dry-run mode first

Before using the workflow live, run a dry-run test against a test user.

---

## How to run

### Queue-based run in Azure Automation

1. Add the user to the group defined in `OffboardingGroupName`.
2. Start `Start-Offboarding-Main` manually or wait for the schedule.
3. Review the job output for each child runbook.
4. When all steps succeed, the user is removed from the queue group automatically.

### Manual dry-run example

```powershell
.\Start-Offboarding-Main.ps1 `
  -OffboardingGroupName "Offboarding-Queue" `
  -ResourceGroupName "your-automation-rg" `
  -AutomationAccountName "your-automation-account" `
  -SubscriptionId "00000000-0000-0000-0000-000000000000" `
  -DryRun `
  -OnlyUserPrincipalName "user@contoso.com"
```

### Test a single step manually

Example:

```powershell
.\Step-01_Immediate-Access-Lockdown.ps1 -UserPrincipalName "user@contoso.com" -DryRun
```

---

## Repository contents

| File | Description |
|---|---|
| `README.md` | Main overview, installation steps, and usage guidance |
| `Start-Offboarding-Main.ps1` | Main orchestrator that queues and runs the workflow |
| `Step-01` to `Step-09` | Individual offboarding actions |
| `AZURE-AUTOMATION-SETUP.md` | Detailed requirements, permissions, modules, runtime notes, and setup checklist |

### Recommended files to include in the repo

For an Azure Repos or Git upload, include at minimum:

- `README.md`
- `Start-Offboarding-Main.ps1`
- `Step-01_Immediate-Access-Lockdown.ps1` through `Step-09_Disable-Account.ps1`
- `AZURE-AUTOMATION-SETUP.md` **(strongly recommended)**

> `AZURE-AUTOMATION-SETUP.md` is not strictly required for the scripts to run, but it is highly recommended for handover, setup, permissions review, and future maintenance.

---

## Safety notes

- Test with **dry-run** first.
- Validate permissions before going live.
- Use a **pilot user** or test account before production use.
- Do not commit tenant-specific secrets, IDs, or private data to the repository.

---

## License

Add a repository `LICENSE` file before public sharing if you want to define reuse terms clearly.

If this project will stay internal in Azure Repos only, a separate license file is optional but still recommended for clarity.
