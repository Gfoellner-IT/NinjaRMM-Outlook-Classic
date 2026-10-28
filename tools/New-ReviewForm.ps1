<#
.SYNOPSIS
  Builds the review dialog forms\ReviewForm.frm (variant B of M8) in a throw-away Excel instance and exports it.

.DESCRIPTION
  A UserForm cannot be written reliably by hand, and Outlook has no way to create one from code. This
  script creates the form with its controls through the VBA designer of Excel, inserts the code behind
  from forms\ReviewForm.code.txt and exports forms\ReviewForm.frm and ReviewForm.frx. The controls are
  named as the code behind expects (lblInfo, lstProposals, btnSelectAll, btnAccept, btnReject, btnClose).
  Captions are set at run time from MatchPromptText, so the exported file is ASCII-only.

  Requires Excel and "Trust access to the VBA project object model" (as Invoke-VbaTests.ps1). Changes no
  security setting and never saves a workbook. Run it again after changing forms\ReviewForm.code.txt.

  If the exported form cannot be imported into Outlook's VBA editor, create it by hand: Insert > UserForm,
  Name = ReviewForm, add the six controls with the names above (lstProposals: ListStyle = option,
  MultiSelect = multi), then paste the code from forms\ReviewForm.code.txt into the code window.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$codeFile = Join-Path $root 'forms\ReviewForm.code.txt'
$target = Join-Path $root 'forms\ReviewForm.frm'

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $project = $wb.VBProject
    try { $null = $project.VBComponents.Count } catch {
        Write-Host 'ENV   No programmatic access to the VBA project in Excel (AccessVBOM = 1 is needed).'
        exit 2
    }

    $comp = $project.VBComponents.Add(3)          # vbext_ct_MSForm
    $comp.Name = 'ReviewForm'
    $comp.Properties.Item('Caption').Value = 'ReviewForm'
    $comp.Properties.Item('Width').Value = 600
    $comp.Properties.Item('Height').Value = 400
    $designer = $comp.Designer

    function Add-Control([string]$progId, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height) {
        $c = $designer.Controls.Add($progId, $name, $true)
        $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
        return $c
    }

    $info = Add-Control 'Forms.Label.1' 'lblInfo' 12 8 570 32
    $info.Caption = 'lblInfo'
    $info.WordWrap = $true

    $list = Add-Control 'Forms.ListBox.1' 'lstProposals' 12 46 570 250
    $list.ListStyle = 1          # fmListStyleOption (check boxes)
    $list.MultiSelect = 1        # fmMultiSelectMulti

    $b1 = Add-Control 'Forms.CommandButton.1' 'btnSelectAll' 12 306 110 28
    $b1.Caption = 'btnSelectAll'
    $b2 = Add-Control 'Forms.CommandButton.1' 'btnAccept' 130 306 200 28
    $b2.Caption = 'btnAccept'
    $b3 = Add-Control 'Forms.CommandButton.1' 'btnReject' 338 306 150 28
    $b3.Caption = 'btnReject'
    $b4 = Add-Control 'Forms.CommandButton.1' 'btnClose' 496 306 86 28
    $b4.Caption = 'btnClose'

    $code = [IO.File]::ReadAllText($codeFile, [Text.Encoding]::GetEncoding(28591))
    $comp.CodeModule.AddFromString($code)

    $comp.Export($target)
    Write-Host "Wrote $target (and ReviewForm.frx)"
}
catch {
    Write-Host "ENV   $($_.Exception.Message)"
    exit 2
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($xl) { try { $xl.Quit() } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
