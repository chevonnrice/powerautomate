# Step 1: Intake flow ("HR Change - 1 Intake")

**What it does:** when a request is submitted, it fills in the Request ID and title, works out the SLA due date (and the
pay period for change type I), logs a comment on the item, pings the People Ops Slack room, and emails the submitter.
This is "Flow 1 — Intake Notification" from your plan.

**Create:** Automated cloud flow → name `HR Change - 1 Intake` → trigger **SharePoint — When an item is created**
* Site Address: your People Ops site · List Name: **HR Change Requests**

## Actions (add in this order, each under the previous one)

### 1. `Settings` (Data Operation → Compose)
```json
{
  "requestPrefix": "HRC-",
  "timeZone": "Eastern Standard Time",
  "payrollCutoffDays": [10, 25]
}
```
`payrollCutoffDays` = cutoff day of month for the 1st–15th period and for the 16th–end period. Use your real cutoff days.

### 2. `Get change type` (SharePoint → Get items)
* List Name: **HR Change Types**
* Filter Query: `Title eq '@{triggerBody()?['ChangeType']?['Value']}'`
* Top Count: `1`

### 3. `Config` (Compose)
```
first(outputs('Get_change_type')?['body/value'])
```

### 4. `Request ID` (Compose)
```
concat(outputs('Settings')?['requestPrefix'], formatNumber(triggerBody()?['ID'], '000000'))
```

### 5. `Submitted Local` (Compose)
```
convertFromUtc(triggerBody()?['Created'], outputs('Settings')?['timeZone'])
```

### 6. `SLA Start` (Compose): a weekend submission starts counting on Monday
```
addDays(startOfDay(outputs('Submitted_Local')), if(equals(dayOfWeek(outputs('Submitted_Local')), 6), 2, if(equals(dayOfWeek(outputs('Submitted_Local')), 0), 1, 0)))
```

### 7. `SLA Days` (Compose)
```
int(coalesce(outputs('Config')?['SLABusinessDays'], 0))
```

### 8. `SLA Due Date` (Compose)
Offboarding (rule *Effective date*) uses the last day or effective date. LOA and *Other* (0 days) get no SLA.
Everything else = submission + N business days.
```
if(equals(outputs('Config')?['DueDateRule'], 'Effective date'), formatDateTime(coalesce(triggerBody()?['LastDay'], triggerBody()?['EffectiveDate']), 'yyyy-MM-dd'), if(equals(outputs('SLA_Days'), 0), null, formatDateTime(addDays(outputs('SLA_Start'), add(add(mul(div(outputs('SLA_Days'), 5), 7), mod(outputs('SLA_Days'), 5)), if(greater(add(dayOfWeek(outputs('SLA_Start')), mod(outputs('SLA_Days'), 5)), 5), 2, 0))), 'yyyy-MM-dd')))
```

### 9–12. Pay period (only used for **I — Extra Shift**, but harmless for the rest)

`Pay Period Date` (Compose):
```
coalesce(triggerBody()?['StartDate'], triggerBody()?['EffectiveDate'])
```
`Pay Period` (Compose):
```
concat(formatDateTime(outputs('Pay_Period_Date'), 'yyyy-MM'), if(lessOrEquals(dayOfMonth(outputs('Pay_Period_Date')), 15), ' (1-15)', ' (16-EOM)'))
```
`Cutoff Date` (Compose):
```
addDays(startOfMonth(outputs('Pay_Period_Date')), sub(if(lessOrEquals(dayOfMonth(outputs('Pay_Period_Date')), 15), first(outputs('Settings')?['payrollCutoffDays']), last(outputs('Settings')?['payrollCutoffDays'])), 1))
```
`After Cutoff` (Compose):
```
and(equals(outputs('Config')?['Code'], 'I'), greater(ticks(outputs('Submitted_Local')), ticks(addDays(outputs('Cutoff_Date'), 1))))
```

### 13. `Update request` (SharePoint → Update item)
* List Name: **HR Change Requests** · Id: **ID** (from the trigger)
* Required columns it asks for: map from the trigger (see the conventions in [README](README.md))

| Field | Value (fx = expression) |
|---|---|
| Title | fx `concat(triggerBody()?['ChangeType']?['Value'], ' — ', triggerBody()?['EmployeeName'])` |
| RequestId | fx `outputs('Request_ID')` |
| SLADueDate | fx `outputs('SLA_Due_Date')` |
| RequestedBy Claims | fx `triggerBody()?['Author']?['Email']` |
| PayPeriod | fx `if(equals(outputs('Config')?['Code'], 'I'), outputs('Pay_Period'), '')` |
| AfterCutoff | fx `outputs('After_Cutoff')` |

> Don't set Status, ApprovalStatus or the task counts here. Their column defaults cover the start, and flows 2–4 own them.

### 14. `Default assignee` (Control → Condition) *(optional)*
* Left: fx `empty(coalesce(outputs('Config')?['DefaultAssigneeEmail'], ''))` · **is equal to** · Right: fx `false`
* **True** branch → SharePoint **Update item** `Set default assignee`: Id = ID, required columns from the trigger,
  **Assignee Claims** = fx `outputs('Config')?['DefaultAssigneeEmail']`

### 15. `Post intake comment` (SharePoint → Send an HTTP request to SharePoint)
This writes to the item's **Comments** pane, the case work log that replaces the Slack List thread.
* Method: `POST`
* Uri: `_api/web/lists/getbytitle('HR Change Requests')/items(@{triggerBody()?['ID']})/Comments`
* Headers:

  | key | value |
  |---|---|
  | `Accept` | `application/json;odata=nometadata` |
  | `Content-Type` | `application/json;odata=nometadata` |
* Body:
```
{ "text": "📥 @{outputs('Request_ID')} received. SLA due: @{coalesce(outputs('SLA_Due_Date'), 'per policy')}@{if(outputs('After_Cutoff'), ' ⚠️ Submitted after the payroll cutoff.', '')}" }
```

### 16. `Ping People Ops` (Slack → Post message (V2))
* Channel: pick your People Ops room
* Message text:
```
:inbox_tray: New HR change request *@{outputs('Request_ID')}* — @{triggerBody()?['ChangeType']?['Value']} — @{triggerBody()?['EmployeeName']}
Effective @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')} • SLA due @{coalesce(outputs('SLA_Due_Date'), 'per policy')}
<@{triggerBody()?['{Link}']}|Open the request>
```

### 17. `Confirm to submitter` (Office 365 Outlook → Send an email (V2))
* To: fx `triggerBody()?['Author']?['Email']`
* Subject: `Received: @{triggerBody()?['ChangeType']?['Value']} for @{triggerBody()?['EmployeeName']} [@{outputs('Request_ID')}]`
* Body: `Thanks, your request @{outputs('Request_ID')} has been received. Target completion: @{coalesce(outputs('SLA_Due_Date'), 'per policy')}.`

## Test

1. Save → **Test → Manually**.
2. Open HR Change Requests → **New** → Change Type **E — Personal Data Change**, fill in the required fields → Save.
3. Within ~1–2 minutes:
   * [ ] Title = `E — Personal Data Change — <name>`, Request ID = `HRC-00000N`, SLA due = the next business day
   * [ ] Comment "📥 HRC-… received" in the item's Comments pane
   * [ ] Slack message in the People Ops room
   * [ ] Confirmation email to you

✅ Then: [Step 2: Checklist router](02-checklist-router.md)
