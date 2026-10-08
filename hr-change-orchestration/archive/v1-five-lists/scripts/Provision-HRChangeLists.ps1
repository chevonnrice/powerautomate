<#
.SYNOPSIS
  Builds the SharePoint side of HR Change + Orchestration: 5 lists, columns, views, and seed data
  (change types, checklist templates A1-I, approval matrix).

.DESCRIPTION
  Re-runnable: existing lists, columns and seeded rows are left alone. Use -ReseedTemplates
  to replace every template row after you edit assets/checklist-templates.json.
  Requires PowerShell 7 + PnP.PowerShell 2.x and an Entra app for PnP interactive login
  (https://pnp.github.io/powershell/articles/registerapplication.html).

.EXAMPLE
  ./Provision-HRChangeLists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/PeopleOps -ClientId <app-id>
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $ClientId,
    [switch] $ReseedTemplates
)

$ErrorActionPreference = 'Stop'
$assets = Join-Path $PSScriptRoot '..' 'assets'
$config      = Get-Content (Join-Path $assets 'sp-list-config.json') -Raw | ConvertFrom-Json
$changeTypes = Get-Content (Join-Path $assets 'change-types.json') -Raw | ConvertFrom-Json
$templates   = Get-Content (Join-Path $assets 'checklist-templates.json') -Raw | ConvertFrom-Json

Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

function ConvertTo-FieldXml($f, $lookupListId) {
    $esc = { param($s) [System.Security.SecurityElement]::Escape([string]$s) }
    $attrs = "Name=""$($f.internalName)"" StaticName=""$($f.internalName)"" DisplayName=""$(& $esc $f.displayName)"""
    if ($f.required) { $attrs += ' Required="TRUE"' }
    if ($f.indexed)  { $attrs += ' Indexed="TRUE"' }
    if ($f.description) { $attrs += " Description=""$(& $esc $f.description)""" }
    $inner = ''
    switch ($f.type) {
        'Text'     { $type = 'Text' }
        'Note'     { $type = 'Note'; $attrs += ' NumLines="6"'; if ($f.richText -eq $false) { $attrs += ' RichText="FALSE"' } }
        'Number'   { $type = 'Number'; if ($null -ne $f.min) { $attrs += " Min=""$($f.min)""" }; if ($null -ne $f.max) { $attrs += " Max=""$($f.max)""" } }
        'Currency' { $type = 'Currency'; $attrs += ' LCID="1033" Decimals="2"' }
        'Boolean'  { $type = 'Boolean' }
        'Date'     { $type = 'DateTime'; $attrs += ' Format="DateOnly"' }
        'User'     { $type = 'User'; $attrs += ' UserSelectionMode="PeopleOnly" UserSelectionScope="0"' }
        'Lookup'   { $type = 'Lookup'; $attrs += " List=""{$lookupListId}"" ShowField=""$($f.lookupField)""" }
        'Choice'   {
            $type = 'Choice'; $attrs += ' Format="Dropdown"'
            $inner = '<CHOICES>' + (($f.choices | ForEach-Object { "<CHOICE>$(& $esc $_)</CHOICE>" }) -join '') + '</CHOICES>'
        }
    }
    if ($null -ne $f.default) { $inner += "<Default>$(& $esc $f.default)</Default>" }
    "<Field Type=""$type"" $attrs>$inner</Field>"
}

$lists = @{}
foreach ($listDef in $config.lists) {
    $list = Get-PnPList -Identity $listDef.title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "Creating list '$($listDef.title)'" -ForegroundColor Cyan
        $list = New-PnPList -Title $listDef.title -Url $listDef.url -Template GenericList -EnableVersioning
        Set-PnPList -Identity $list -Description $listDef.description | Out-Null
    }
    if ($listDef.itemLevelPermissions) {
        # Submitters see and edit only their own requests. HR (site owners / Full Control) see everything.
        Set-PnPList -Identity $list -ReadSecurity 2 -WriteSecurity 2 | Out-Null
    }
    $lists[$listDef.title] = $list

    foreach ($f in $listDef.fields) {
        if ($f.internalName -eq 'Title') {
            Set-PnPField -List $list -Identity 'Title' -Values @{ Title = $f.displayName; Required = $false } | Out-Null
            continue
        }
        if (Get-PnPField -List $list -Identity $f.internalName -ErrorAction SilentlyContinue) { continue }
        $lookupId = if ($f.type -eq 'Lookup') { $lists[$f.lookupList].Id } else { $null }
        Write-Host "  + $($f.internalName)"
        Add-PnPFieldFromXml -List $list -FieldXml (ConvertTo-FieldXml $f $lookupId) | Out-Null
        if ($f.group -ne 'Case' -or $f.internalName -in 'Status', 'Assignee', 'SLADueDate') {
            Add-PnPViewField -List $list -Identity 'All Items' -Field $f.internalName -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

# ---------------------------------------------------------------- seed data
function Seed($listTitle, $rows, [switch] $Force) {
    $list = $lists[$listTitle]
    $existing = Get-PnPListItem -List $list -PageSize 500
    if ($existing.Count -gt 0 -and -not $Force) { Write-Host "'$listTitle' already has rows - not seeding"; return }
    if ($Force) { $existing | ForEach-Object { Remove-PnPListItem -List $list -Identity $_.Id -Force } }
    foreach ($r in $rows) {
        $values = @{}
        $r.PSObject.Properties | Where-Object { $null -ne $_.Value -and $_.Value -ne '' } | ForEach-Object { $values[$_.Name] = $_.Value }
        Add-PnPListItem -List $list -Values $values | Out-Null
    }
    Write-Host "Seeded $($rows.Count) rows into '$listTitle'" -ForegroundColor Green
}

Seed 'HR Change Types' $changeTypes.changeTypes
Seed 'HR Approval Matrix' $changeTypes.approvalMatrix

$templateRows = foreach ($p in $templates.templates.PSObject.Properties) {
    $i = 0
    foreach ($step in $p.Value) {
        $i++
        [pscustomobject]@{
            Title                = $step.title
            ChangeType           = $p.Name
            StepOrder            = $i
            Details              = $step.details
            OwnerRole            = if ($step.owner) { $step.owner } else { 'HR' }
            Automation           = if ($step.automation) { $step.automation } else { 'None' }
            BlockedUntilApproved = [bool]$step.blocked
            DueOnEffectiveDate   = [bool]$step.dueOnEffectiveDate
            Active               = $true
        }
    }
}
Seed 'HR Change Templates' $templateRows -Force:$ReseedTemplates

# ---------------------------------------------------------------- views
function Ensure-View($listTitle, $name, $fields, $query) {
    $list = $lists[$listTitle]
    if (-not (Get-PnPView -List $list -Identity $name -ErrorAction SilentlyContinue)) {
        Write-Host "View '$listTitle / $name'"
        Add-PnPView -List $list -Title $name -Fields $fields -Query $query | Out-Null
    }
}
$open = "<Neq><FieldRef Name='Status'/><Value Type='Choice'>3-Resolved</Value></Neq>"
$caseFields = 'RequestId', 'Title', 'Status', 'Assignee', 'SLADueDate', 'ApprovalStatus', 'ProgressPct', 'EffectiveDate'
Ensure-View 'HR Change Requests' 'Open by SLA' $caseFields "<Where>$open</Where><OrderBy><FieldRef Name='SLADueDate'/></OrderBy>"
Ensure-View 'HR Change Requests' 'My cases' $caseFields "<Where><And>$open<Eq><FieldRef Name='Assignee'/><Value Type='Integer'><UserID/></Value></Eq></And></Where><OrderBy><FieldRef Name='SLADueDate'/></OrderBy>"
Ensure-View 'HR Change Requests' '0-New (pick me up)' $caseFields "<Where><Eq><FieldRef Name='Status'/><Value Type='Choice'>0-New</Value></Eq></Where><OrderBy><FieldRef Name='Created'/></OrderBy>"
Ensure-View 'HR Change Requests' 'Awaiting approval' ($caseFields + 'CurrentApprover') "<Where><Eq><FieldRef Name='ApprovalStatus'/><Value Type='Choice'>Pending</Value></Eq></Where>"
Ensure-View 'HR Change Requests' 'By change type' $caseFields "<GroupBy Collapse='TRUE'><FieldRef Name='ChangeType'/></GroupBy><Where>$open</Where>"
Ensure-View 'HR Change Tasks' 'Open tasks by request' 'StepOrder', 'Title', 'TaskStatus', 'OwnerRole', 'DueDate', 'Details' `
    "<GroupBy Collapse='FALSE'><FieldRef Name='RequestId'/></GroupBy><Where><And><Neq><FieldRef Name='TaskStatus'/><Value Type='Choice'>Done</Value></Neq><Neq><FieldRef Name='TaskStatus'/><Value Type='Choice'>Cancelled</Value></Neq></And></Where><OrderBy><FieldRef Name='StepOrder'/></OrderBy>"
Ensure-View 'HR Change Templates' 'By change type' 'StepOrder', 'Title', 'OwnerRole', 'Automation', 'BlockedUntilApproved', 'Active' `
    "<GroupBy Collapse='TRUE'><FieldRef Name='ChangeType'/></GroupBy><OrderBy><FieldRef Name='StepOrder'/></OrderBy>"

$web = Get-PnPWeb
Write-Host "`nDone." -ForegroundColor Green
Write-Host "Front door (pin this in #lane-people-talent): $($web.Url.TrimEnd('/'))/$($config.lists[0].url)/NewForm.aspx"
Write-Host "Next: add the conditional show/hide formulas from assets/form-conditional-formulas.md to the New form."
