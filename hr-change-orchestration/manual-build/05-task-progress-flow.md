# Step 5: Task progress ("HR Change - 5 Task Progress")

**What it does:** whenever a task changes, it recounts that request's tasks and updates **TasksDone** and **ProgressPct**
on the request (the progress bar the plan wanted from native subtasks). When the last task is ticked, it comments and pings
the owner that the case is ready to resolve. It never auto-resolves, so LOA can stay open.

**Create:** Automated cloud flow → `HR Change - 5 Task Progress` → **SharePoint — When an item is created or modified** → list **HR Change Tasks**

**Trigger → Settings:**
* **Concurrency control: On, Degree of parallelism: 1** (so two quick ticks can't overwrite each other's count)
* **Trigger conditions** (two lines). The first skips newly created tasks, because the router already set the total:
  ```
  @not(equals(triggerBody()?['Created'], triggerBody()?['Modified']))
  ```
  ```
  @not(empty(triggerBody()?['Request']?['Id']))
  ```

## Actions

1. **`Get sibling tasks`** (Get items) → HR Change Tasks · Filter Query `RequestId eq '@{triggerBody()?['RequestId']}'` · Top Count `500`
2. **`Done tasks`** (Filter array). From fx `outputs('Get_sibling_tasks')?['body/value']`, advanced:
   `@or(equals(item()?['TaskStatus']?['Value'], 'Done'), equals(item()?['TaskStatus']?['Value'], 'N/A'))`
3. **`Cancelled tasks`** (Filter array). Same From, advanced: `@equals(item()?['TaskStatus']?['Value'], 'Cancelled')`
4. **`Total`** (Compose): `sub(length(outputs('Get_sibling_tasks')?['body/value']), length(body('Cancelled_tasks')))`
5. **`Done`** (Compose): `length(body('Done_tasks'))`
6. **`Get request`** (Get item) → HR Change Requests · Id fx `triggerBody()?['Request']?['Id']`
7. **`Update progress`** (Update item) → HR Change Requests · Id fx `triggerBody()?['Request']?['Id']` · required columns from **Get request** (not the trigger):
   * TasksTotal fx `outputs('Total')` · TasksDone fx `outputs('Done')`
   * ProgressPct fx `if(equals(outputs('Total'), 0), 0, div(mul(outputs('Done'), 100), outputs('Total')))`
8. **`Just completed?`** (Condition). Left:
   ```
   and(equals(outputs('Done'), outputs('Total')), greater(outputs('Total'), 0), less(int(coalesce(outputs('Get_request')?['body/ProgressPct'], 0)), 100), not(equals(outputs('Get_request')?['body/Status']?['Value'], '3-Resolved')))
   ```
   is equal to fx `true` → **True**:
   * **Send an HTTP request to SharePoint** (set up like Step 1 #15, but the Uri uses
     `items(@{triggerBody()?['Request']?['Id']})`). Body:
     `{ "text": "☑️ All checklist tasks are complete — set Status to 3-Resolved to close the case." }`
   * Slack **Post message (V2)** → People Ops room:
     `:ballot_box_with_check: All tasks done on *@{triggerBody()?['RequestId']}* — @{coalesce(outputs('Get_request')?['body/Assignee']?['DisplayName'], 'owner')}, ready to resolve.`

## Test
Open any request's tasks and tick them *Done* one by one.
* [ ] Progress % on the request goes up after each tick (allow about a minute)
* [ ] After the last one: the "☑️ All checklist tasks are complete" comment and Slack ping, and the Status is **unchanged**

✅ Then: [Step 6: Daily digest](06-daily-digest-flow.md)
