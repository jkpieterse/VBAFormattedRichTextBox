$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'Treeview 26-5.xlsm'
$target = Join-Path $PSScriptRoot 'RichText Demo 1.0.xlsm'
if (Test-Path $target) { Remove-Item $target -Force }
Copy-Item $source $target
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, 0, $false)
    $vb = $wb.VBProject
    $workbookModule = $vb.VBComponents.Item('ThisWorkbook').CodeModule
    if ($workbookModule.CountOfLines -gt 0) {
        $workbookModule.DeleteLines(1, $workbookModule.CountOfLines)
    }
    foreach ($name in @('ufDemo','UserForm1','ufRichTextDemo','clsNode','clsTreeView','modAddIcons','modDemo','modDocNav')) {
        try { $vb.VBComponents.Remove($vb.VBComponents.Item($name)) } catch {}
    }
    foreach ($file in @('clsRichTextCharacters.cls','clsRichTextBox.cls')) {
        $vb.VBComponents.Import((Join-Path $PSScriptRoot $file)) | Out-Null
    }
    $module = $vb.VBComponents.Import((Join-Path $PSScriptRoot 'modRichTextDemo.bas'))

    $hasFormExport = Test-Path (Join-Path $PSScriptRoot 'ufRichTextDemo.frm')
    if ($hasFormExport) {
        $form = $vb.VBComponents.Import((Join-Path $PSScriptRoot 'ufRichTextDemo.frm'))
    }
    else {
        $form = $vb.VBComponents.Add(3)
        $form.Name = 'ufRichTextDemo'
    $form.Properties.Item('Caption').Value = 'RichText Demo 1.0'
    $form.Properties.Item('Width').Value = 560
    $form.Properties.Item('Height').Value = 350
    $form.Properties.Item('StartUpPosition').Value = 1
    $frame = $form.Designer.Controls.Add('Forms.Frame.1', 'frRichText', $true)
    $frame.Left = 12; $frame.Top = 12; $frame.Width = 365; $frame.Height = 250; $frame.Caption = 'Formatted text'; $frame.ScrollBars = 0
    $label = $form.Designer.Controls.Add('Forms.Label.1', 'labInfo', $true)
    $label.Left = 390; $label.Top = 16; $label.Width = 145; $label.Height = 64; $label.WordWrap = $true; $label.Caption = 'RichText status'
    foreach ($spec in @(@('cmdEdit','Edit text',390,92),@('cmdSource','Edit markup',390,126),@('cmdCommit','Commit',390,160),@('cmdCancel','Cancel',390,194),@('cmdReset','Reset sample',390,228))) {
        $button = $form.Designer.Controls.Add('Forms.CommandButton.1', $spec[0], $true)
        $button.Caption = $spec[1]; $button.Left = $spec[2]; $button.Top = $spec[3]; $button.Width = 145; $button.Height = 23
    }
        $form.CodeModule.AddFromString(@'
Private WithEvents mcRichText As clsRichTextBox
Private msAppName As String

Public Property Let AppName(ByVal sValue As String)
    msAppName = sValue
End Property

Private Sub UserForm_Initialize()
    Set mcRichText = New clsRichTextBox
    Set mcRichText.HostFrame = Me.frRichText
    mcRichText.Markup = "This is <b>bold</b>, <i>italic</i>, <u>underlined</u>, and <span style=""color:#1F5F92"">colored</span>."
    UpdateInfo
End Sub

Private Sub cmdEdit_Click()
    mcRichText.BeginEdit
End Sub

Private Sub cmdSource_Click()
    mcRichText.BeginSourceEdit
End Sub

Private Sub cmdCommit_Click()
    mcRichText.CommitEdit
    UpdateInfo
End Sub

Private Sub cmdCancel_Click()
    mcRichText.CancelEdit
    UpdateInfo
End Sub

Private Sub cmdReset_Click()
    mcRichText.Markup = "This is <b>bold</b>, <i>italic</i>, <u>underlined</u>, and <span style=""color:#1F5F92"">colored</span>."
    UpdateInfo
End Sub

Private Sub mcRichText_Change()
    UpdateInfo
End Sub

Private Sub mcRichText_EditCommitted(ByVal SourceMode As Boolean)
    UpdateInfo
End Sub

Private Sub mcRichText_ParseError(ByVal Message As String)
    Me.labInfo.Caption = "Markup error: " & Message
End Sub

Private Sub frRichText_Resize()
    If Not mcRichText Is Nothing Then mcRichText.Resize
End Sub

Private Sub UpdateInfo()
    If mcRichText Is Nothing Then Exit Sub
    Me.labInfo.Caption = "Text length: " & Len(mcRichText.Text) & vbCrLf & _
        "Markup: " & mcRichText.Markup
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mcRichText Is Nothing Then mcRichText.Terminate
    Set mcRichText = Nothing
End Sub
'@)
    }
    if (-not $hasFormExport) {
        $form.Designer.BackColor = 15987699
        $form.Designer.ForeColor = 4210752
        foreach ($control in $form.Designer.Controls) {
            $control.BackColor = 15987699
            $control.ForeColor = 4210752
        }
        foreach ($name in @('cmdEdit','cmdSource','cmdCommit','cmdCancel','cmdReset')) {
            $button = $form.Designer.Controls.Item($name)
            $button.BackColor = 15332846
            $button.ForeColor = -2147483641
        }
    }

    function Get-OrCreateSheet([string]$name) {
        try { return $wb.Worksheets.Item($name) } catch {}
        $sheet = $wb.Worksheets.Add($wb.Worksheets.Item($wb.Worksheets.Count))
        $sheet.Name = $name
        return $sheet
    }

    function Reset-Sheet($sheet) {
        $sheet.Cells.Clear() | Out-Null
        for ($i = $sheet.Shapes.Count; $i -ge 1; $i--) { $sheet.Shapes.Item($i).Delete() }
        $sheet.Columns.Item('A').ColumnWidth = 28
        $sheet.Columns.Item('B').ColumnWidth = 34
        $sheet.Columns.Item('C').ColumnWidth = 90
        $sheet.Columns.Item('C').WrapText = $true
        $sheet.Rows.Item(1).Font.Bold = $true
        $sheet.Rows.Item(1).Font.Size = 16
        $sheet.Rows.Item(3).Font.Bold = $true
        $sheet.Range('A:C').VerticalAlignment = -4160
    }

    function Write-Table($sheet, [int]$startRow, [object[]]$rows) {
        $rowIndex = $startRow
        foreach ($row in $rows) {
            for ($columnIndex = 0; $columnIndex -lt $row.Count; $columnIndex++) {
                $sheet.Cells.Item($rowIndex, $columnIndex + 1).Value = $row[$columnIndex]
            }
            $rowIndex++
        }
        $header = $sheet.Range($sheet.Cells.Item($startRow, 1), $sheet.Cells.Item($startRow, 3))
        $header.Font.Bold = $true
        $header.Interior.Color = 14277081
    }

    $desiredSheets = @('Overview','clsRichTextBox','clsRichTextCharacters','Host UserForm','Versions')
    foreach ($sheetName in $desiredSheets) {
        $sheet = Get-OrCreateSheet $sheetName
        Reset-Sheet $sheet
    }
    foreach ($sheet in @($wb.Worksheets)) {
        if ($desiredSheets -notcontains $sheet.Name) { $sheet.Delete() | Out-Null }
    }
    for ($i = $desiredSheets.Count - 1; $i -ge 0; $i--) {
        $wb.Worksheets.Item($desiredSheets[$i]).Move($wb.Worksheets.Item(1)) | Out-Null
    }

    $overview = $wb.Worksheets.Item('Overview')
    $overview.Range('A1').Value = 'RichText Demo 1.0'
    $overview.Range('A3').Value = 'All-VBA MSForms rich text box overview'
    $overview.Range('A5').Value = 'This workbook demonstrates clsRichTextBox, a reusable control class that renders formatted text inside a design-time MSForms.Frame by creating dynamic Label controls for contiguous formatting runs and a temporary TextBox while editing.'
    $overview.Range('A7').Value = 'The public entry point is modRichTextDemo.DemoNow, which opens ufRichTextDemo. The form shows the required host wiring and optional buttons for plain-text editing, markup editing, commit, cancel, and reset.'
    $button = $overview.Buttons().Add(12, 156, 150, 24)
    $button.Caption = 'Open RichText demo'
    $button.OnAction = 'DemoNow'

    $box = $wb.Worksheets.Item('clsRichTextBox')
    $box.Range('A1').Value = 'clsRichTextBox'
    $box.Range('A3').Value = 'Reusable Rich Text Box control object'
    $box.Range('A5').Value = 'Create this class in a UserForm, assign HostFrame to an existing MSForms.Frame, then set Text or Markup. The class owns rendering, parsing, keyboard focus, editing, resizing, and cleanup.'
    Write-Table $box 7 @(
        @('Member','Type','Description'),
        @('HostFrame','Property','Required host MSForms.Frame. Assigning it captures default fonts and colors, creates the hidden focus proxy, and renders the current content.'),
        @('Text','Property','Plain text represented by the control. Setting it replaces the content and clears existing character formatting.'),
        @('Markup','Property','Canonical HTML-like source for the formatted text. Setting it parses supported tags and keeps the last valid content if parsing fails.'),
        @('ReadOnly','Property','Disables user editing when True while still allowing display and programmatic content changes.'),
        @('DefaultFontName / DefaultFontSize','Properties','Override the font inherited from the host Frame for unformatted characters.'),
        @('DefaultForeColor / DefaultBackColor','Properties','Override the colors inherited from the host Frame for unformatted characters.'),
        @('Characters(start, length)','Function','Returns a one-based clsRichTextCharacters range for programmatic formatting, similar to Excel Characters.'),
        @('BeginEdit / BeginSourceEdit','Methods','Start plain-text editing or markup-source editing using the temporary TextBox.'),
        @('ToggleEditMode','Method','Switch between plain-text and markup editing while preserving edit state.'),
        @('CommitEdit / CancelEdit','Methods','Accept or discard the current TextBox edit.'),
        @('HandleKeyDown','Method','Lets the host forward F2, Ctrl+Enter, Escape, and markup/text toggle keys.'),
        @('SetFocus / EnterExit / Resize / Terminate','Methods','Host lifecycle methods for focus indication, keyboard ownership, dynamic sizing, and generated-control cleanup.'),
        @('Change / EditCommitted / EditCancelled / ParseError','Events','Notify the host when content changes, editing finishes or cancels, or markup cannot be parsed.')
    )

    $chars = $wb.Worksheets.Item('clsRichTextCharacters')
    $chars.Range('A1').Value = 'clsRichTextCharacters'
    $chars.Range('A3').Value = 'One-based formatting range object'
    $chars.Range('A5').Value = 'Characters returns this temporary object for a start position and length. Each formatting property reads a single value when the whole range matches, returns Null for mixed values, and writes the value to every character in the range.'
    Write-Table $chars 7 @(
        @('Property or method','Value type','Description'),
        @('FontName','Variant/String','Gets or sets the font name for the selected character range.'),
        @('FontSize','Variant/Numeric','Gets or sets the font size for the selected character range.'),
        @('Bold','Variant/Boolean','Gets or sets bold formatting for the selected character range.'),
        @('Italic','Variant/Boolean','Gets or sets italic formatting for the selected character range.'),
        @('Underline','Variant/Boolean','Gets or sets underline formatting for the selected character range.'),
        @('Strikethrough','Variant/Boolean','Gets or sets strikethrough formatting for the selected character range.'),
        @('ForeColor','Variant/Long','Gets or sets the foreground color for the selected character range.'),
        @('BackColor','Variant/Long','Gets or sets the background color for the selected character range.'),
        @('Clear','Method','Removes custom formatting from the selected character range and returns it to the control defaults.')
    )

    $hostSheet = $wb.Worksheets.Item('Host UserForm')
    $hostSheet.Range('A1').Value = 'Host UserForm'
    $hostSheet.Range('A3').Value = 'Required host objects and routines'
    $hostSheet.Range('A5').Value = 'The host form must contain a design-time MSForms.Frame, named frRichText in this demo. The frame is the surface where clsRichTextBox creates labels, the edit TextBox, and the hidden focus proxy.'
    Write-Table $hostSheet 7 @(
        @('Host item','Required?','Purpose'),
        @('Private WithEvents mcRichText As clsRichTextBox','Yes','Keeps the control instance alive and exposes its events to the UserForm.'),
        @('UserForm_Initialize','Yes','Creates clsRichTextBox, assigns HostFrame, and loads initial Text or Markup.'),
        @('UserForm_Activate','Yes','Calls SetFocus so keyboard editing shortcuts are routed to the RichText control.'),
        @('frRichText_Enter / frRichText_Exit','Yes','Forward focus transitions to EnterExit so the control can show or hide the focus border.'),
        @('frRichText_Resize','Yes','Calls Resize so generated labels and edit controls stay aligned to the host Frame.'),
        @('UserForm_QueryClose','Yes','Calls Terminate and releases the class before the form unloads.'),
        @('cmdEdit / cmdSource / cmdCommit / cmdCancel / cmdReset','No','Demo-only buttons that expose plain-text edit, markup edit, commit, cancel, and sample reset commands.'),
        @('labInfo and lblLink','No','Demo-only status and documentation link controls; they are not required when embedding the class in another form.')
    )

    $ver = $wb.Worksheets.Item('Versions')
    $ver.Range('A1').Value = 'RichText Demo 1.0'
    $ver.Range('A3').Value = 'Version history'
    Write-Table $ver 5 @(
        @('Version','Date','Notes'),
        @('1.0','16-Sep-2026','Standalone RichText demonstration workbook with RichText-specific overview, class reference, host wiring notes, and properly labeled demo launcher button.')
    )
    foreach ($name in @('ufRichTextDemo','modRichTextDemo','clsRichTextCharacters','clsRichTextBox')) {
        $extension = if ($name -eq 'ufRichTextDemo') { '.frm' } elseif ($name -eq 'modRichTextDemo') { '.bas' } else { '.cls' }
        $vb.VBComponents.Item($name).Export((Join-Path $PSScriptRoot ($name + $extension)))
    }
    $wb.Saved = $false
    $wb.Save()
}
finally {
    if ($null -ne $wb) { $wb.Close($true) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output 'RichText build completed.'
Get-ChildItem $PSScriptRoot | Select-Object Name,Length
