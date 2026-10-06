<#
.SYNOPSIS
  Creates the "HR Tickets" and "HR Triage Log" SharePoint lists from assets/sp-list-config.json.

.DESCRIPTION
  Safe to re-run: existing lists and columns are kept. Requires PnP.PowerShell 2.x
  (Install-Module PnP.PowerShell -Scope CurrentUser) and an Entra app registration for PnP
  interactive login (see https://pnp.github.io/powershell/articles/registerapplication.html).

.EXAMPLE
  ./Provision-HRTriageLists.ps1 -SiteUrl https://contoso.sharepoint.com/sites/HR -ClientId <app-id>
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $ClientId,
    [string] $ConfigPath = (Join-Path $PSScriptRoot '..' 'assets' 'sp-list-config.json')
)

$ErrorActionPreference = 'Stop'
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

foreach ($listDef in $config.lists) {
    $list = Get-PnPList -Identity $listDef.title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "Creating list '$($listDef.title)'"
        $list = New-PnPList -Title $listDef.title -Template GenericList -EnableVersioning -OnQuickLaunch
        Set-PnPList -Identity $list -Description $listDef.description | Out-Null
    } else {
        Write-Host "List '$($listDef.title)' already exists"
    }

    foreach ($f in $listDef.fields) {
        if ($f.internalName -eq 'Title') { continue }

        $existing = Get-PnPField -List $list -Identity $f.internalName -ErrorAction SilentlyContinue
        if (-not $existing) {
            Write-Host "  + $($f.internalName) ($($f.type))"
            $common = @{ List = $list; DisplayName = $f.internalName; InternalName = $f.internalName; AddToDefaultView = $true }
            switch ($f.type) {
                'Choice'   { $existing = Add-PnPField @common -Type Choice -Choices $f.choices }
                'Note'     { $existing = Add-PnPField @common -Type Note }
                'User'     { $existing = Add-PnPField @common -Type User }
                'Boolean'  { $existing = Add-PnPField @common -Type Boolean }
                'Number'   { $existing = Add-PnPField @common -Type Number }
                'DateTime' { $existing = Add-PnPField @common -Type DateTime }
                default    { $existing = Add-PnPField @common -Type Text }
            }
        }

        $values = @{}
        if ($f.description) { $values['Description'] = $f.description }
        if ($null -ne $f.default) { $values['DefaultValue'] = [string]$f.default }
        if ($f.indexed) { $values['Indexed'] = $true }
        if ($f.type -eq 'Note' -and $f.richText -eq $false) { $values['RichText'] = $false }
        if ($f.type -eq 'Number') {
            if ($null -ne $f.min) { $values['MinimumValue'] = $f.min }
            if ($null -ne $f.max) { $values['MaximumValue'] = $f.max }
        }
        if ($values.Count) { Set-PnPField -List $list -Identity $f.internalName -Values $values | Out-Null }
    }

    # The flows filter on these columns, so index them.
    if ($listDef.title -eq 'HR Tickets') {
        Set-PnPField -List $list -Identity 'Title' -Values @{ Indexed = $true } | Out-Null
    }
}

# Helpful views for the HR team
$tickets = Get-PnPList -Identity 'HR Tickets'
$viewFields = 'TicketNumber', 'Title', 'Category', 'Priority', 'Status', 'AssignedTo', 'DueDate', 'LastActivity'
$views = @{
    'My open tickets' = "<Where><And><Eq><FieldRef Name='AssignedTo'/><Value Type='Integer'><UserID/></Value></Eq><And><Neq><FieldRef Name='Status'/><Value Type='Choice'>Resolved</Value></Neq><Neq><FieldRef Name='Status'/><Value Type='Choice'>Closed</Value></Neq></And></And></Where><OrderBy><FieldRef Name='DueDate'/></OrderBy>"
    'Needs triage'    = "<Where><Eq><FieldRef Name='Status'/><Value Type='Choice'>Needs Triage</Value></Eq></Where><OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy>"
    'All open'        = "<Where><And><Neq><FieldRef Name='Status'/><Value Type='Choice'>Resolved</Value></Neq><Neq><FieldRef Name='Status'/><Value Type='Choice'>Closed</Value></Neq></And></Where><OrderBy><FieldRef Name='DueDate'/></OrderBy>"
}
foreach ($name in $views.Keys) {
    if (-not (Get-PnPView -List $tickets -Identity $name -ErrorAction SilentlyContinue)) {
        Write-Host "Creating view '$name'"
        Add-PnPView -List $tickets -Title $name -Fields $viewFields -Query $views[$name] | Out-Null
    }
}

$log = Get-PnPList -Identity 'HR Triage Log'
if (-not (Get-PnPView -List $log -Identity 'Needs verdict' -ErrorAction SilentlyContinue)) {
    Add-PnPView -List $log -Title 'Needs verdict' -Fields 'Title', 'Sender', 'Decision', 'Reason', 'AIConfidence', 'HumanVerdict' `
        -Query "<Where><IsNull><FieldRef Name='HumanVerdict'/></IsNull></Where><OrderBy><FieldRef Name='AIConfidence'/></OrderBy>" | Out-Null
}

Write-Host "`nDone. List ids (for the hrt_TicketsList / hrt_TriageLogList environment variables):"
'HR Tickets', 'HR Triage Log' | ForEach-Object { Get-PnPList -Identity $_ } | Select-Object Title, Id | Format-Table
Write-Host "Reminder: for sensitive tickets, restrict who can open the HR Tickets list, or move sensitive tickets to a separate locked-down list."
