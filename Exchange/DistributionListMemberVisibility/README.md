# Distribution list member visibility

Reports every distribution list in Exchange Online together with its members, showing which members are **hidden from the address lists**, and optionally unhides them.

A hidden member still receives mail sent to the list — hiding only removes them from the global address list and every other address list. That is why a distribution list can look perfectly healthy while the people on it cannot be found in the address book by anyone trying to reach them directly.

| Script | Purpose |
|--------|---------|
| [`ReportDistributionListMemberVisibility.ps1`](ReportDistributionListMemberVisibility.ps1) | Reports list membership and per-member address-list visibility. Unhides hidden members when `-Unhide` is supplied. |

## Why this sits in `Exchange/` and not `Reports/`

The repository keeps reporting scripts in `Reports/`. This one lives here because `-Unhide` writes to the tenant: it is an action script that produces a report, not a report that happens to have a switch. The read-only run is the default, and nothing changes without `-Unhide` or `-UnhideGroups`.

## Usage

```powershell
# Read-only. Every distribution group and mail-enabled security group.
.\ReportDistributionListMemberVisibility.ps1

# Check the output format first, against a random sample of 10 lists.
.\ReportDistributionListMemberVisibility.ps1 -TestMode

# Widen to dynamic distribution groups and Microsoft 365 Groups too.
.\ReportDistributionListMemberVisibility.ps1 -GroupType All -ReportPath C:\Reports

# Preview exactly what would be unhidden, changing nothing.
.\ReportDistributionListMemberVisibility.ps1 -Unhide -WhatIf

# Unhide every hidden member of one list, after confirmation.
.\ReportDistributionListMemberVisibility.ps1 -Group "All Staff" -Unhide

# Unhide every hidden member in scope, and the hidden lists themselves, unattended.
.\ReportDistributionListMemberVisibility.ps1 -Unhide -UnhideGroups -Force
```

## Choosing which members to unhide

Rarely is the answer "all of them". Four ways to narrow it, all of which apply to the **change** only — the report still covers every list and member in scope.

```powershell
# Just these, by address, alias, display name or Entra object id.
.\ReportDistributionListMemberVisibility.ps1 -Unhide -Member ada@contoso.com,grace@contoso.com

# Everything except these. The usual shape of the request.
.\ReportDistributionListMemberVisibility.ps1 -Unhide -ExcludeMember ceo@contoso.com,chair@contoso.com

# Wildcards, so a whole partner domain can be named at once.
.\ReportDistributionListMemberVisibility.ps1 -Unhide -Member *@fabrikam.com

# Pick from a numbered list of what was actually found.
.\ReportDistributionListMemberVisibility.ps1 -Group "All Staff" -Unhide -SelectMembers
```

`-Member` and `-ExcludeMember` combine: `-Member` selects the pool, then `-ExcludeMember` removes from it. `-SelectMembers` is interactive and cannot be combined with either.

`-SelectMembers` prints what it found and accepts `all`, `1,4,7`, ranges like `2-6`, and exclusions like `all -3,-9`:

```
Hidden members in scope:
   1. Ada Lovelace  <ada@contoso.com>  [UserMailbox, via Exchange Online]
   2. Ext Partner  <partner@fabrikam.com>  [GuestMailUser, via Graph (guest)]
   3. Legacy Payroll  <payroll@contoso.com>  [MailContact, via on-premises AD]

Selection: all -2
```

Three things worth knowing:

- Members you deselect are recorded in the report with a **Not selected** action, not dropped. The report stays an honest account of what was found versus what was asked for, so a hidden member left hidden has a stated reason.
- A `-Member` or `-ExcludeMember` value matching nothing is a warning, not a failure — a typo'd address does not take the run down, but it also does not pass silently.
- These switches only affect `-Unhide`, so the script refuses them on a read-only run rather than accepting them and doing nothing. To narrow the *report*, use `-Group`. And they do not apply to `-UnhideGroups`, which is scoped by `-Group`.

Because `-SelectMembers` is its own confirmation, it skips the batch prompt that otherwise follows.

Run `Get-Help .\ReportDistributionListMemberVisibility.ps1 -Full` for every parameter.

## Output

Two files, timestamped, written to `-ReportPath` (the current folder by default):

- **CSV** — one row per group-and-member pair. This is the finest grain the report is built from; every roll-up in the HTML is computed from it.
- **HTML** — a self-contained report. Summary cards, a hidden-members-by-recipient-type chart, a list-centric table with expandable membership, and a member-centric table filtered to hidden members on open. It makes no external requests (the chart is inline SVG, there is no CDN reference), so it opens on an isolated machine, and it carries no branding.

## Three ways a member gets unhidden

Not every recipient is unhidden the same way, and the report's **Change via** column says which route each member needs. This is the part that catches people out.

| Change via | Applies to | What actually runs |
|-----------|-----------|-------------------|
| **Exchange Online** | Mailboxes, mail users, mail contacts, groups, public folders | The `Set-*` cmdlet for that recipient type — `Set-Mailbox`, `Set-MailUser`, `Set-DistributionGroup`, and so on. Exchange splits this property by recipient type; there is no single cmdlet for it. |
| **Microsoft Graph (guest)** | Entra ID guest users (`GuestMailUser`) | `Update-MgUser -UserId <entra object id> -ShowInAddressList:$true`. A guest's visibility is the `showInAddressList` property on the **Entra ID user object**, not an Exchange setting, so `Set-MailUser` is the wrong tool. The script connects to Graph only when there is a hidden guest to change; `-SkipGraph` leaves guests alone. |
| **On-premises AD** | Directory-synchronized objects | Nothing, in the cloud. Exchange Online refuses the write, so the change has to be made in Active Directory (`msExchHideFromAddressLists`) and synchronized. These are skipped and listed as a work list unless `-IncludeDirSynced` is given. |

### Connect order matters: Graph before Exchange

`ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` ship **different, incompatible** copies of `Microsoft.Identity.Client` and `Microsoft.IdentityModel.Abstractions`. Only one version of each loads per process, so the first module to *authenticate* wins and the second one's sign-in fails with:

```
Method not found: '!0 Microsoft.Identity.Client.BaseAbstractApplicationBuilder`1
.WithLogging(Microsoft.IdentityModel.Abstractions.IIdentityLogger, Boolean)'
```

**Updating the modules does not fix this.** Verified against the current releases:

| Module | `Microsoft.Identity.Client` | `IdentityModel.Abstractions` |
|---|---|---|
| ExchangeOnlineManagement 3.10.1 | 4.83.1.0 | 8.19.2.26155 |
| Microsoft.Graph.Authentication 2.41.2 | 4.90.0.0 | 8.18.0.26120 |

So when `-Unhide` could touch a guest, the script authenticates to **Graph first**, before Exchange Online. Two consequences:

- An `-Unhide` run without `-SkipGraph` asks for a Graph sign-in up front, even if there turn out to be no hidden guests. Pass `-SkipGraph` when you know there are none.
- If you connected to Exchange Online yourself before starting the script, the Graph half is already lost for that session. The run still completes — guests are reported as needing Graph, and the companion script below is written.

### The companion guest script

Whenever hidden guests are left unchanged — no `-Unhide`, `-SkipGraph`, or a failed Graph connection — the script writes `<report name>_GuestRemediation.ps1` beside the report. It carries the affected guests as data and does the Graph work before touching Exchange:

```powershell
Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome
Update-MgUser   -UserId <id> -ShowInAddressList:$true     # per guest
Connect-ExchangeOnline -ShowBanner:$false
Set-MailUser    -Identity <id> -HiddenFromAddressListsEnabled $false
```

Run it **in a fresh session with no Exchange Online connection open**. It supports `-WhatIf`.

### Guests are hidden by default

Guests invited through the invitation manager have `showInAddressList` set to `false` at invitation time. So guests routinely show up in this report as hidden without anyone having hidden them — that is Entra ID's default, not a misconfiguration someone introduced.

Updating `showInAddressList` also updates the Exchange hide-from-address-lists setting, but the two are stored separately and **where they disagree, the Exchange setting wins**. After the Graph call this script therefore also clears `HiddenFromAddressListsEnabled` on the mail user if Exchange still has it set, so the two agree and the next person to look does not find a contradiction.

### Directory-synchronized objects that never change

If an on-premises change to `msExchHideFromAddressLists` never appears in the cloud, check that the object's `mailNickname` (alias) is populated. The Entra Connect sync rule for Exchange attributes is scoped on that attribute being non-null, so an empty alias silently skips the whole Exchange attribute set.

## Permissions

| Operation | Needs |
|-----------|-------|
| Report (default) | **View-Only Recipients**. `ExchangeOnlineManagement` module. No Graph needed. |
| `-Unhide` / `-UnhideGroups` | **Mail Recipients** or **Recipient Management**. |
| Unhiding guest users | **User Administrator**, or the `User.ReadWrite.All` Graph scope. `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users` modules. |

The Graph modules are checked for at the point they are needed, not up front, so a read-only run — and a run with no hidden guests in it — needs neither installed.

## Safety

- Read-only by default. `-Unhide` and `-UnhideGroups` are the only switches that change anything.
- `-WhatIf` and `-Confirm` are supported. `-WhatIf` names the exact command it would run against each object, including the Graph call for guests.
- Before any batch, the script prints the plan — how many objects, split by change route, how many were selected versus left hidden, and what will be skipped and why — and asks for confirmation. `-Force` skips that prompt.
- `-Member`, `-ExcludeMember` and `-SelectMembers` narrow the change to specific recipients. See [Choosing which members to unhide](#choosing-which-members-to-unhide).
- Each unique member is changed once, however many lists it appears in.
- `-TestMode` cannot be combined with `-Unhide`: remediating a random sample would leave the tenant half-changed.
- Hiding a list and hiding its members are independent decisions, often deliberate for different reasons, which is why unhiding the lists themselves needs its own switch.

## What this does not cover

- **Address book policies.** A member visible in the GAL can still be invisible to a given user through an ABP. Visibility here means `HiddenFromAddressListsEnabled` only.
- **`HiddenFromExchangeClientsEnabled`** on Microsoft 365 Groups, which hides a group from Outlook independently of the address lists. `Get-Recipient` does not return it, so it is neither reported nor changed. Check it with `Get-UnifiedGroup` if a group stays hidden in Outlook after being unhidden here.
- Members that are not mail-enabled recipients, which cannot appear in distribution list membership in the first place.

## References

- [Manage address lists in Exchange Online](https://learn.microsoft.com/exchange/address-books/address-lists/manage-address-lists)
- [user resource type — `showInAddressList`](https://learn.microsoft.com/graph/api/resources/user)
- [Cloud-based management of Exchange attributes in hybrid environments](https://learn.microsoft.com/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management)
- [Changes to `msExchHideFromAddressLists` not reflected in Exchange Online](https://learn.microsoft.com/troubleshoot/exchange/administration/changes-to-msexchangehiddenfromaddresslist-not-updated-against-recipient)
