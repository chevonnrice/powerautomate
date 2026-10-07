# Step 2: Checklist router ("HR Change - 2 Checklist Router")

**What it does:** reads the change type, finds that type's steps in **HR Change Templates**, and creates one row per
step in **HR Change Tasks**, linked to the request. That gives you real subtasks instead of a checklist message.
Types with no template (*Other*) get a "needs manual scoping" comment instead of silently falling through.

It's a separate flow from Intake on purpose: as your plan says, a failure in one shouldn't take down the other.

**Create:** Automated cloud flow → `HR Change - 2 Checklist Router` → **SharePoint — When an item is created** → list **HR Change Requests**

## Actions

### 1. `Settings` (Compose)
```json
{ "requestPrefix": "HRC-" }
```

### 2. `Request ID` (Compose)
```
concat(outputs('Settings')?['requestPrefix'], formatNumber(triggerBody()?['ID'], '000000'))
```
(The same formula as Intake, so both flows agree on the ID without waiting for each other.)

### 3. `Get change type` (SharePoint → Get items) → list **HR Change Types**
* Filter Query: `Title eq '@{triggerBody()?['ChangeType']?['Value']}'` · Top Count `1`

### 4. `Config` (Compose)
```
first(outputs('Get_change_type')?['body/value'])
```

### 5. `Needs Approval` (Compose)
```
greater(length(concat(coalesce(outputs('Config')?['ApprovalStage1'], ''), coalesce(outputs('Config')?['ApprovalStage2'], ''), coalesce(outputs('Config')?['ApprovalStage3'], ''))), 0)
```

### 6. `Get template steps` (SharePoint → Get items) → list **HR Change Templates**
* Filter Query: `ChangeType eq '@{triggerBody()?['ChangeType']?['Value']}' and Active eq 1`
* Order By: `StepOrder asc` · Top Count: `200`

### 7. `Has template` (Condition)
* Left: fx `length(outputs('Get_template_steps')?['body/value'])` · **is greater than** · Right: `0`

#### True branch

**7a. `Each step`** (Control → Apply to each). Select an output: fx `outputs('Get_template_steps')?['body/value']`

Inside it, **`Create task`** (SharePoint → Create item) → list **HR Change Tasks**:

| Field | Value |
|---|---|
| Title | fx `items('Each_step')?['Title']` |
| Request Id *(the lookup)* | fx `triggerBody()?['ID']` |
| RequestId | fx `outputs('Request_ID')` |
| ChangeType | fx `triggerBody()?['ChangeType']?['Value']` |
| StepOrder | fx `items('Each_step')?['StepOrder']` |
| Details | fx `coalesce(items('Each_step')?['Details'], '')` |
| OwnerRole | fx `coalesce(items('Each_step')?['OwnerRole'], 'HR')` |
| Automation | fx `coalesce(items('Each_step')?['Automation'], 'None')` |
| TaskStatus Value | fx `if(and(equals(items('Each_step')?['BlockedUntilApproved'], true), outputs('Needs_Approval')), 'Blocked - awaiting approval', 'Not started')` |
| DueDate | fx `if(equals(items('Each_step')?['DueOnEffectiveDate'], true), coalesce(triggerBody()?['LastDay'], triggerBody()?['EffectiveDate']), null)` |

**7b. `Set task totals`** (SharePoint → Update item), placed *after* the loop (not inside it) → **HR Change Requests**, Id = trigger **ID**, required columns from the trigger:
* TasksTotal: fx `length(outputs('Get_template_steps')?['body/value'])` · TasksDone: `0` · ProgressPct: `0`

**7c. `Post checklist comment`** (Send an HTTP request to SharePoint): set up like Step 1 #15 (POST, same Uri and headers). Body:
```
{ "text": "📋 @{outputs('Config')?['Code']} checklist created — @{length(outputs('Get_template_steps')?['body/value'])} tasks in HR Change Tasks." }
```

#### False branch

**7d. `Post no-template comment`** (Send an HTTP request to SharePoint): same setup. Body:
```
{ "text": "⚠️ No checklist template for this change type — needs manual scoping." }
```
**7e. `Ping needs scoping`** (Slack → Post message (V2)) → People Ops room:
`:warning: @{outputs('Request_ID')} (@{triggerBody()?['ChangeType']?['Value']}) has no checklist template — needs manual scoping.`

## Test

Submit another **E — Personal Data Change** request (or delete and recreate the one from Step 1).
* [ ] 5 rows in **HR Change Tasks** with the request's ID, Step 1–5, *Not started*
* [ ] Request shows Tasks Total = 5
* [ ] Comment "📋 E checklist created — 5 tasks"
* [ ] An **Other** request → the "needs manual scoping" comment + Slack ping

> Tip: on HR Change Tasks, create a view grouped by `RequestId` and sorted by `StepOrder`. That's your per-case checklist.

✅ Then: [Step 3: Resolution](03-resolution-flow.md)
