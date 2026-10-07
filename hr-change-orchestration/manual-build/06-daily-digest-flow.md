# Step 6: Daily digest + payroll cutoff reminder ("HR Change - 6 Daily")

**What it does:**
* **Weekdays:** posts an SLA digest to the People Ops room listing overdue requests, requests due today, and 0-New requests nobody has picked up.
* **Before each payroll cutoff** (your "nice to add"): N business days ahead, asks the managers channel whether there are
  extra shifts to report, with a link to the form.

**Create:** **Scheduled cloud flow** → `HR Change - 6 Daily` → repeat every **1 Day**, at **8:00**, in your time zone.

## Actions

### 1. `Settings` (Compose)
```json
{
  "timeZone": "Eastern Standard Time",
  "payrollCutoffDays": [10, 25],
  "reminderBusinessDaysBefore": 2,
  "formUrl": "https://<tenant>.sharepoint.com/sites/<site>/Lists/<HR Change Requests>/NewForm.aspx"
}
```
Get `formUrl` by opening the list → **+ New** and copying the browser address.

### 2. `Today` (Compose)
```
convertFromUtc(utcNow(), outputs('Settings')?['timeZone'])
```

### 3. SLA digest

**3a. `Get open requests`** (Get items) → HR Change Requests · Filter Query `Status ne '3-Resolved'` · Order By `SLADueDate asc` · Top Count `5000`

**3b. Three Filter arrays**, each From fx `outputs('Get_open_requests')?['body/value']`, advanced mode:
* `Overdue`: `@and(not(empty(item()?['SLADueDate'])), less(formatDateTime(item()?['SLADueDate'], 'yyyyMMdd'), formatDateTime(outputs('Today'), 'yyyyMMdd')))`
* `Due today`: `@and(not(empty(item()?['SLADueDate'])), equals(formatDateTime(item()?['SLADueDate'], 'yyyyMMdd'), formatDateTime(outputs('Today'), 'yyyyMMdd')))`
* `Not picked up`: `@equals(item()?['Status']?['Value'], '0-New')`

**3c. Three Selects** (Data Operation → Select). From = the matching filter's body, e.g. fx `body('Overdue')`.
Switch **Map** to text mode (the small **T** icon) and paste:
```
concat('• <', item()?['{Link}'], '|', item()?['RequestId'], '> ', item()?['ChangeType']?['Value'], ' — ', item()?['EmployeeName'], ' — ', coalesce(item()?['Assignee']?['DisplayName'], 'unassigned'))
```
Name them `Overdue lines`, `Due today lines`, `Not picked up lines`.

**3d. `Weekday with news?`** (Condition). Left:
```
and(greater(dayOfWeek(outputs('Today')), 0), less(dayOfWeek(outputs('Today')), 6), greater(add(add(length(body('Overdue')), length(body('Due_today'))), length(body('Not_picked_up'))), 0))
```
is equal to fx `true` → **True**: Slack **Post message (V2)** → People Ops room:
```
:bar_chart: *HR Change — daily SLA check* (@{length(outputs('Get_open_requests')?['body/value'])} open)
*:red_circle: Overdue (@{length(body('Overdue'))})*
@{join(body('Overdue_lines'), decodeUriComponent('%0A'))}
*:large_yellow_circle: Due today (@{length(body('Due_today'))})*
@{join(body('Due_today_lines'), decodeUriComponent('%0A'))}
*:white_circle: 0-New, not picked up (@{length(body('Not_picked_up'))})*
@{join(body('Not_picked_up_lines'), decodeUriComponent('%0A'))}
```

### 4. Payroll cutoff reminder

**4a. `Cutoffs this month`** (Select). From fx `outputs('Settings')?['payrollCutoffDays']`, Map in text mode:
```
formatDateTime(addDays(startOfMonth(outputs('Today')), sub(int(item()), 1)), 'yyyy-MM-dd')
```
**4b. `Cutoffs next month`** (Select). Same From, Map:
```
formatDateTime(addDays(startOfMonth(addToTime(outputs('Today'), 1, 'Month')), sub(int(item()), 1)), 'yyyy-MM-dd')
```
**4c. `Remind today`** (Filter array). From fx `union(body('Cutoffs_this_month'), body('Cutoffs_next_month'))`, advanced:
```
@equals(formatDateTime(outputs('Today'), 'yyyy-MM-dd'), formatDateTime(addDays(addDays(item(), mul(int(outputs('Settings')?['reminderBusinessDaysBefore']), -1)), if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['reminderBusinessDaysBefore']), -1))), 6), -1, if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['reminderBusinessDaysBefore']), -1))), 0), -2, 0))), 'yyyy-MM-dd'))
```
(It counts back N days from each cutoff. If that lands on a weekend, the reminder moves to the Friday before.)

**4d. `Reminder due?`** (Condition): Left fx `length(body('Remind_today'))` · is greater than · `0` → **True**: Slack **Post message (V2)** → **managers channel**:
```
:alarm_clock: *Payroll cutoff is @{formatDateTime(first(body('Remind_today')), 'dddd, MMM d')}.* Any extra shifts or fractional FTE pickups to report for this pay period? Submit them as *I — Fractional FTE / Extra Shift Pickup*: <@{outputs('Settings')?['formUrl']}|HR change request form>
```

## Test
* **Test → Manually → Run.** If you have an open request with a past SLA date (edit one to test), the digest posts.
* To test the reminder without waiting, temporarily set `payrollCutoffDays` to a date 2 business days from today, run it, then set it back.

🎉 That's the whole build. Remaining Phase 2 items (Jira, Word letters, Monday) are in the main [README](../README.md#whats-next-phase-2-items-not-built-yet-blocked-on-your-dependencies-table).
