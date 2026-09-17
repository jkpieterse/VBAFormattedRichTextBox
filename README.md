# VBAFormattedRichTextBox

(c) 2026, Copyright JKP Application Development Services, all rights reserved

An all-VBA rich text box look-alike for MSForms `UserForm`s. It renders formatted
text with dynamic `Label` controls and uses a temporary `TextBox` for editing, so it
does not require an external RichEdit or ActiveX control.

The control supports per-character formatting, plain-text editing, and HTML-like
markup editing inside an `MSForms.Frame` supplied by the host form.

For background and a longer explanation, see the accompanying article:
[vba-rich-text-box-control.aspx](vba-rich-text-box-control.aspx)

## Files

| File | Purpose |
|---|---|
| `clsRichTextBox.cls` | The reusable rich text control. Owns parsing, rendering, editing, keyboard handling, resizing, and cleanup. |
| `clsRichTextCharacters.cls` | An Excel-`Characters`-like object for reading and writing formatting on a one-based character range. |
| `ufRichTextDemo.frm` and `ufRichTextDemo.frx` | Exported demo `UserForm` showing the required host wiring and editing commands. |
| `modRichTextDemo.bas` | Demo entry point. Run `DemoNow` to open the demonstration form. |
| `RichText Demo 1.0.xlsm` | Ready-to-run demonstration workbook with documentation sheets. |
| `References.txt` | VBA references required by the source modules. |
| `build-richtext.ps1` | Optional script for rebuilding the demo workbook from an Excel template. |

## Getting started

1. Import `clsRichTextCharacters.cls` and `clsRichTextBox.cls` into your VBA project.
2. Add an `MSForms.Frame` to a `UserForm`.
3. Create the control and attach it to the frame:

```vb
Private WithEvents mcRichText As clsRichTextBox

Private Sub UserForm_Initialize()
    Set mcRichText = New clsRichTextBox
    Set mcRichText.HostFrame = Me.frRichText
    mcRichText.Markup = "This is <b>bold</b> and <i>italic</i>."
End Sub
```

4. Forward the host frame's resize and focus lifecycle events:

```vb
Private Sub frRichText_Resize()
    If Not mcRichText Is Nothing Then mcRichText.Resize
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mcRichText Is Nothing Then mcRichText.Terminate
    Set mcRichText = Nothing
End Sub
```

The exported `ufRichTextDemo.frm` contains the complete host implementation,
including focus forwarding and optional demo buttons.

## Formatting characters

`Characters(Start, Length)` returns a temporary formatting range. Formatting
properties can be read or assigned across the selected characters:

```vb
With mcRichText.Characters(1, 5)
    .Bold = True
    .ForeColor = vbRed
End With

mcRichText.Characters(1, 5).Clear
```

When a range contains mixed values, a formatting property returns `Null`.

## Supported markup

The parser supports the following HTML-like tags:

- `<b>` and `<strong>`
- `<i>` and `<em>`
- `<u>`
- `<s>`
- `<font color="#RRGGBB">`
- `<span style="color:#RRGGBB;background-color:#RRGGBB;font-size:12">`
- `<size=12>`

Setting `Markup` parses the source into per-character formatting. Setting `Text`
replaces the content with unformatted plain text. Invalid markup raises `ParseError`
and preserves the last valid content.

## Key members

- **Content**: `Text`, `Markup`, `ReadOnly`
- **Appearance**: `DefaultFontName`, `DefaultFontSize`, `DefaultForeColor`, `DefaultBackColor`
- **Editing**: `BeginEdit`, `BeginSourceEdit`, `ToggleEditMode`, `CommitEdit`, `CancelEdit`
- **Lifecycle**: `SetFocus`, `EnterExit`, `Resize`, `Terminate`
- **Events**: `Change`, `EditCommitted`, `EditCancelled`, `ParseError`

## Keyboard editing

- `F2` starts plain-text editing when the host forwards the key to `HandleKeyDown`.
- `Ctrl+Enter` commits an edit.
- `Escape` cancels an edit.
- `Ctrl+?` or `Ctrl+<` toggles between text and markup editing where supported by
  the keyboard layout.
- Hosts can call `ToggleEditMode` directly for a layout-independent fallback.

## Demo

Open `RichText Demo 1.0.xlsm` and run `modRichTextDemo.DemoNow`. The workbook
includes an overview, class reference sheets, host wiring notes, and a form with
plain-text editing, markup editing, commit, cancel, and reset commands.

## Requirements

- Windows desktop Excel with MSForms 2.0
- Trust access to the VBA project object model when importing modules or rebuilding
  the demo workbook

The optional `build-richtext.ps1` script also requires PowerShell, Excel, and a
`Treeview 26-5.xlsm` template in the parent folder. The committed demo workbook
means the build script is not needed to use the control.

## Limitations

- Designed for `UserForm`/MSForms hosts only; it does not provide a worksheet or
  task-pane control.
- Text is rendered with real controls, so extremely long documents may be slower
  than a native rich text editor.
- The repository currently contains source and demo files but no open-source
  license. All rights are reserved unless the copyright holder grants permission.