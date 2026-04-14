# Azure Automation Setup Requirements

This document lists the **permissions**, **variables**, **modules**, **runtime settings**, and **other required configuration** for the offboarding solution built in this workspace.

> **Public release note:** all names, domains, IDs, and group labels below are sample values only. Replace them with the values from your own tenant before use.

## Runbooks in scope

- `Start-Offboarding-Main`
- `Step-01_Immediate-Access-Lockdown`
- `Step-02_Convert-Mailbox-To-Shared`
- `Step-03_Remove-User-Groups`
- `Step-04_Remove-Distribution-And-Delegation`
- `Step-05_Remove-SharePoint-Access`
- `Step-06_Remove-Calendar-And-Room-Access`
- `Step-07_Cleanup-User-Profile`
- `Step-08_Remove-Remaining-Licenses`
- `Step-09_Disable-Account`

---

## 1) Required Azure Automation Variables

> Use the **canonical names below**. Some scripts support aliases, but these are the recommended names to keep configured.

| Variable name | Required | Example value | Used by | Purpose |
|---|---|---:|---|---|
| `OffboardingGroupName` | Yes | `Offboarding-Queue` | `Start-Offboarding-Main`, `Step-03` | Queue group containing users to offboard |
| `ResourceGroupName` | Yes | `your-automation-rg` | `Start-Offboarding-Main` | Resource group that hosts the Automation Account |
| `AutomationAccountName` | Yes | `your-automation-account` | `Start-Offboarding-Main` | Allows the main runbook to start and monitor child runbooks |
| `SubscriptionId` | Yes | `00000000-0000-0000-0000-000000000000` | `Start-Offboarding-Main` | Ensures the runbooks use the correct Azure context |
| `ExchangeOrganization` | Yes | `contoso.onmicrosoft.com` | `Step-02`, `Step-04`, `Step-06` | Exchange Online organization for app-only/managed identity connection |
| `SharePointAdminUrl` | Yes | `https://contoso-admin.sharepoint.com` | `Step-05` | SharePoint admin endpoint used by `PnP.PowerShell` |
| `OffboardingStepDelaySeconds` | Optional | `30` | `Start-Offboarding-Main` | Minimum delay between steps for propagation |
| `OffboardingPollSeconds` | Optional | `15` | `Start-Offboarding-Main` | Polling interval while waiting for child jobs |
| `OffboardingMaxStepWaitSeconds` | Optional | `3600` | `Start-Offboarding-Main` | Maximum time to wait for a child runbook before timeout |

---

## 2) Required Microsoft Graph **Application Permissions** for the Automation Account Managed Identity

These are granted to the **system-assigned managed identity** of the Automation Account.

| Permission | Required | Used by | Why it is needed |
|---|---|---|---|
| `User.Read.All` | Yes | Main, 01, 07, 08, 09 | Look up users and read account state |
| `User.ReadWrite.All` | Yes | 07, 09 | Update user profile fields and disable account |
| `User-PasswordProfile.ReadWrite.All` | Yes | 01 | Reset the user password |
| `User.RevokeSessions.All` | Yes | 01 | Revoke current sign-in sessions |
| `UserAuthenticationMethod.ReadWrite.All` | Yes | 01 | Remove MFA / Authenticator / phone / FIDO / WHfB methods |
| `User-Phone.ReadWrite.All` | Yes | 07 | Clear `businessPhones` and `mobilePhone` |
| `Group.ReadWrite.All` | Yes | Main, 03 | Read and update group objects |
| `GroupMember.ReadWrite.All` | Yes | Main, 03 | Remove users from groups and from the offboarding queue group |
| `Directory.ReadWrite.All` | Recommended / Yes | 03, 07 | Directory updates and broader Graph directory operations |
| `RoleManagement.ReadWrite.Directory` | Required when role-assignable groups exist | 03 | Remove users from **role-assignable** groups such as `usermailbox-sharedmailbox` |
| `Calendars.ReadWrite` | Yes | 06 | Graph fallback for calendar cleanup |

### Notes
- The current clean 9-step design **does not require** Intune wipe permissions because the legacy device-wipe step is not part of this rebuilt flow.
- If you reintroduce device wipe or OAuth cleanup later, additional permissions will be needed.

---

## 3) Required Entra ID / Microsoft 365 **Admin Roles** for the Managed Identity

These roles are recommended for the supported public workflow.

| Role | Required | Why it is needed |
|---|---|---|
| `User Administrator` | Yes | Update core Entra user properties and disable the account |
| `Exchange Administrator` | Yes | Connect to Exchange Online and run mailbox / distribution / calendar actions |
| `SharePoint Administrator` | Yes | Remove SharePoint access in `Step-05` |
| `Privileged Authentication Administrator` | Yes | Clear phone-related auth/profile values and protected auth data |
| `Privileged Role Administrator` | Required when users belong to role-assignable groups | Needed for removing memberships from role-assignable groups in `Step-03` |

---

## 4) Required **Service-Specific** App Permissions Outside Graph

| Service | Permission | Required | Used by | Notes |
|---|---|---|---|---|
| SharePoint Online | `Sites.FullControl.All` | Yes | `Step-05` | Required for `PnP.PowerShell` managed-identity access to enumerate and remove site users |

---

## 5) Required Azure RBAC / Access to Azure Resources

| Scope | Recommended role | Why it is needed |
|---|---|---|
| Automation Account or hosting Resource Group | `Contributor` or `Automation Contributor` | Lets `Start-Offboarding-Main` start child runbooks and read job state/output |
| Subscription / RG for read access | `Reader` as needed | Helpful for consistent Az context and diagnostics |
| Automation Account managed identity | `Automation Contributor` or `Contributor` on the Automation Account RG | Required so the managed identity can resolve subscription context and start runbooks with `Connect-AzAccount -Identity` |

> The key operational requirement is that the managed identity can successfully run `Start-AzAutomationRunbook`, `Get-AzAutomationJob`, and related Az.Automation calls inside the same Automation Account.

---

## 6) Required Modules and Runtime Environments

| Runbook(s) | Runtime | Required modules/packages | Notes |
|---|---|---|---|
| `Start-Offboarding-Main` | `PowerShell` | `Az.Accounts`, `Az.Automation` | Starts and monitors child runbooks |
| `Step-01`, `Step-03`, `Step-07`, `Step-08`, `Step-09` | `PowerShell` | `Az.Accounts` | Use Graph REST via `Get-AzAccessToken` |
| `Step-02`, `Step-04`, `Step-06` | `PowerShell` | `Az.Accounts`, `ExchangeOnlineManagement` | Use Exchange Online app-only / managed identity patterns |
| `Step-05_Remove-SharePoint-Access` | `PowerShell` | `PnP.PowerShell`, `Az.Accounts` | Works as a classic PowerShell runbook when `PnP.PowerShell` is imported into the Automation account; optionally use a dedicated custom runtime for better isolation |

### Important runtime note for Step 05
`Step-05_Remove-SharePoint-Access` can be linked to a custom runtime if needed, but the default tested deployment uses classic Azure Automation `PowerShell` runbooks for all steps.

If you do choose a dedicated runtime for Step 05, link it to:

- **Runbook type:** `PowerShell`
- **Runtime environment:** `PowerShell-74-Custom`

This is helpful because `PnP.PowerShell` is most reliable in a dedicated PowerShell 7.4 runtime for Azure Automation scenarios.

---

## 7) Other Required Configuration

| Item | Required | Recommended / Example value | Purpose |
|---|---|---|---|
| System-assigned managed identity | Yes | Enabled on your Automation Account | Used for Graph, Exchange, SharePoint, and Az access |
| Offboarding queue group | Yes | `Offboarding-Queue` | Users added here are picked up by `Start-Offboarding-Main` |
| Main schedule | Recommended | `Offboarding-Every-12-Hours` | Runs the orchestrator every 12 hours |
| Child runbook publishing | Yes | All 9 step runbooks published | The main runbook resolves and starts them by name |
| Step delays between child jobs | Yes | Built into `Start-Offboarding-Main` | Allows propagation before the next step starts |
| Max child wait timeout | Recommended | `3600` seconds | Prevents the main runbook from appearing stuck forever |

---

## 8) Recommended Execution Order

| Order | Runbook | Purpose |
|---:|---|---|
| 1 | `Step-01_Immediate-Access-Lockdown` | Revoke sessions, reset password, remove auth methods |
| 2 | `Step-02_Convert-Mailbox-To-Shared` | Convert mailbox before downstream cleanup |
| 3 | `Step-03_Remove-User-Groups` | Remove Entra/M365 groups except the protected offboarding queue |
| 4 | `Step-04_Remove-Distribution-And-Delegation` | Remove Exchange distribution/mail-enabled memberships and shared mailbox delegation |
| 5 | `Step-05_Remove-SharePoint-Access` | Remove SharePoint direct site access |
| 6 | `Step-06_Remove-Calendar-And-Room-Access` | Remove meeting/calendar/resource access |
| 7 | `Step-07_Cleanup-User-Profile` | Remove manager and clear non-essential profile fields |
| 8 | `Step-08_Remove-Remaining-Licenses` | Remove direct licenses and report group-based assignments |
| 9 | `Step-09_Disable-Account` | Disable the Entra account as the final step |

---

## 9) Operational Notes

- Add users to the group configured in `OffboardingGroupName` to queue them for offboarding.
- `Start-Offboarding-Main` runs the steps **one by one**, waits for each child job to finish, and pauses between steps for propagation.
- In a **live run**, users are removed from the queue group only after all steps succeed.
- In a **dry-run**, no destructive changes are made and the user remains in the queue group.

---

## 10) Quick Setup Checklist

- [ ] Enable **system-assigned managed identity** on the Automation Account
- [ ] Grant the Graph **application permissions** listed above
- [ ] Grant the Entra **admin roles** listed above
- [ ] Grant SharePoint Online app permission **`Sites.FullControl.All`**
- [ ] Import required modules (`Az.Accounts`, `Az.Automation`, `ExchangeOnlineManagement`, `PnP.PowerShell`)
- [ ] Ensure `Step-05_Remove-SharePoint-Access` is linked to **`PowerShell-74-Custom`**
- [ ] Create the required Automation Variables
- [ ] Publish `Start-Offboarding-Main` and `Step-01` through `Step-09`
- [ ] Create/enable the **12-hour schedule**
- [ ] Add users to the group configured in `OffboardingGroupName` to begin offboarding
