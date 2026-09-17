VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRichTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : clsRichTextBox
' Purpose: All-VBA rich-text display, markup parsing, editing, and host integration.
'---------------------------------------------------------------------------------------
Option Explicit

Private Const MC_FORMAT_COUNT As Long = 10
Private Const MC_FONT_NAME As Long = 0
Private Const MC_FONT_SIZE As Long = 1
Private Const MC_BOLD As Long = 2
Private Const MC_ITALIC As Long = 3
Private Const MC_UNDERLINE As Long = 4
Private Const MC_STRIKETHROUGH As Long = 5
Private Const MC_WEIGHT As Long = 6
Private Const MC_CHARSET As Long = 7
Private Const MC_FORE_COLOR As Long = 8
Private Const MC_BACK_COLOR As Long = 9

Private moHost As MSForms.Frame
Private WithEvents moEditBox As MSForms.TextBox
Attribute moEditBox.VB_VarHelpID = -1
Private WithEvents moFocusProxy As MSForms.TextBox
Attribute moFocusProxy.VB_VarHelpID = -1
Private mlTextLength As Long
Private msText As String
Private mvFormats() As Variant
Private mcolControlNames As Collection
Private msEditText As String
Private mbSourceMode As Boolean
Private mbReadOnly As Boolean
Private mbUpdating As Boolean
Private msDefaultFontName As String
Private msDefaultFontSize As Single
Private mlDefaultForeColor As Long
Private mlDefaultBackColor As Long
Private mlHostBorderColor As Long
Private mlHostBorderStyle As Long

Public Event Change()
Public Event EditCommitted(ByVal SourceMode As Boolean)
Public Event EditCancelled()
Public Event ParseError(ByVal Message As String)

Private Const MC_KEY_CTRL As Long = 2
Private Const MC_KEY_SHIFT As Long = 1

'=======================================================================================
' HOST INTEGRATION: assign the design-time MSForms.Frame that owns generated controls.
'=======================================================================================
Public Property Set HostFrame(ByVal oFrame As MSForms.Frame)
    Set moHost = oFrame
    msDefaultFontName = oFrame.Font.Name
    msDefaultFontSize = oFrame.Font.Size
    mlDefaultForeColor = oFrame.ForeColor
    mlDefaultBackColor = oFrame.BackColor
    mlHostBorderColor = oFrame.BorderColor
    mlHostBorderStyle = oFrame.BorderStyle
    EnsureFocusProxy
    Render
End Property

' Return the currently assigned host Frame.
Public Property Get HostFrame() As MSForms.Frame
    Set HostFrame = moHost
End Property

' Give the hidden keyboard proxy focus so host forms can receive F2.
Public Sub SetFocus()
    EnsureFocusProxy
    If moEditBox Is Nothing Then
        On Error Resume Next
        moFocusProxy.SetFocus
        On Error GoTo 0
    End If
    ShowFocusBorder
End Sub

' Show or restore the host focus indicator from the UserForm Enter/Exit events.
Public Sub EnterExit(ByVal bExit As Boolean)
    If bExit Then
        HideFocusBorder
    Else
        ShowFocusBorder
    End If
End Sub

' Return the plain text represented by the control.
Public Property Get Text() As String
    Text = msText
End Property

' Replace the text and clear its existing character formatting.
Public Property Let Text(ByVal sValue As String)
    ResetText sValue
    Render
    RaiseEvent Change
End Property

' Return canonical markup generated from the current text and formats.
Public Property Get Markup() As String
    Markup = ToMarkup
End Property

' Parse markup into text and per-character formatting.
Public Property Let Markup(ByVal sValue As String)
    Dim sPreviousMarkup As String
    sPreviousMarkup = ToMarkup
    If Not ParseMarkup(sValue) Then
        ParseMarkup sPreviousMarkup
        Exit Property
    End If
    Render
    RaiseEvent Change
End Property

' Return or set whether editing is disabled.
Public Property Get ReadOnly() As Boolean
    ReadOnly = mbReadOnly
End Property

' Enable or disable user editing.
Public Property Let ReadOnly(ByVal bValue As Boolean)
    mbReadOnly = bValue
End Property

' Return or set the default font name used by unformatted characters.
Public Property Get DefaultFontName() As String
    DefaultFontName = msDefaultFontName
End Property

' Override the inherited default font name.
Public Property Let DefaultFontName(ByVal sValue As String)
    msDefaultFontName = sValue
    Render
End Property

' Return or set the default font size used by unformatted characters.
Public Property Get DefaultFontSize() As Single
    DefaultFontSize = msDefaultFontSize
End Property

' Override the inherited default font size.
Public Property Let DefaultFontSize(ByVal sValue As Single)
    msDefaultFontSize = sValue
    Render
End Property

' Return or set the default foreground color.
Public Property Get DefaultForeColor() As Long
    DefaultForeColor = mlDefaultForeColor
End Property

' Override the inherited default foreground color.
Public Property Let DefaultForeColor(ByVal lValue As Long)
    mlDefaultForeColor = lValue
    Render
End Property

' Return or set the default background color.
Public Property Get DefaultBackColor() As Long
    DefaultBackColor = mlDefaultBackColor
End Property

' Override the inherited default background color.
Public Property Let DefaultBackColor(ByVal lValue As Long)
    mlDefaultBackColor = lValue
    Render
End Property

' Initialize safe defaults before a host Frame is assigned.
Private Sub Class_Initialize()
    msDefaultFontName = "Calibri"
    msDefaultFontSize = 10
    mlDefaultForeColor = vbWindowText
    mlDefaultBackColor = vbWhite
    ResetText vbNullString
End Sub

' Return a one-based character range for programmatic formatting.
Public Function Characters(Optional ByVal lStart As Long = 1, Optional ByVal lLength As Long = -1) As clsRichTextCharacters
    Dim oCharacters As clsRichTextCharacters
    ValidateRange lStart, lLength
    If lLength = -1 Then lLength = mlTextLength - lStart + 1
    Set oCharacters = New clsRichTextCharacters
    oCharacters.Initialize Me, lStart, lLength
    Set Characters = oCharacters
End Function

' Read one formatting property across a range; mixed values return Null.
Friend Function FormatValue(ByVal lStart As Long, ByVal lLength As Long, ByVal sProperty As String) As Variant
    Dim lPosition As Long
    Dim lProperty As Long
    Dim vFirst As Variant
    ValidateRange lStart, lLength
    lProperty = FormatPropertyIndex(sProperty)
    vFirst = CharacterFormatValue(lStart, lProperty)
    For lPosition = lStart + 1 To lStart + lLength - 1
        If Not ValuesEqual(vFirst, CharacterFormatValue(lPosition, lProperty)) Then
            FormatValue = Null
            Exit Function
        End If
    Next lPosition
    FormatValue = vFirst
End Function

' Apply one formatting property to every character in a range.
Friend Sub SetFormatValue(ByVal lStart As Long, ByVal lLength As Long, ByVal sProperty As String, ByVal vValue As Variant)
    Dim lPosition As Long
    Dim lProperty As Long
    Dim vFormat As Variant
    ValidateRange lStart, lLength
    lProperty = FormatPropertyIndex(sProperty)
    For lPosition = lStart To lStart + lLength - 1
        vFormat = CharacterFormat(lPosition)
        vFormat(lProperty) = vValue
        mvFormats(lPosition) = vFormat
    Next lPosition
    Render
    RaiseEvent Change
End Sub

' Remove custom formatting from every character in a range.
Friend Sub ClearFormat(ByVal lStart As Long, ByVal lLength As Long)
    Dim lPosition As Long
    ValidateRange lStart, lLength
    For lPosition = lStart To lStart + lLength - 1
        mvFormats(lPosition) = Empty
    Next lPosition
    Render
    RaiseEvent Change
End Sub

' Start plain-text editing in the temporary TextBox.
Public Sub BeginEdit()
    BeginEditInternal False
End Sub

' Start source/markup editing in the temporary TextBox.
Public Sub BeginSourceEdit()
    BeginEditInternal True
End Sub

' Toggle the active editor between plain text and canonical markup.
Public Sub ToggleEditMode()
    Dim lSelectionStart As Long
    Dim lSelectionLength As Long
    Dim sPreviousMarkup As String

    If moEditBox Is Nothing Then
        BeginEditInternal False
        Exit Sub
    End If

    lSelectionStart = moEditBox.SelStart
    lSelectionLength = moEditBox.SelLength
    If mbSourceMode Then
        sPreviousMarkup = ToMarkup
        If Not ParseMarkup(moEditBox.Text) Then
            ParseMarkup sPreviousMarkup
            Exit Sub
        End If
        mbSourceMode = False
        moEditBox.Text = msText
    Else
        ApplyPlainTextEdit moEditBox.Text
        mbSourceMode = True
        moEditBox.Text = ToMarkup
    End If
    moEditBox.SelStart = MinLong(lSelectionStart, Len(moEditBox.Text))
    moEditBox.SelLength = MinLong(lSelectionLength, Len(moEditBox.Text) - moEditBox.SelStart)
End Sub

' Handle host-forwarded keys when the editor itself does not have focus.
Public Sub HandleKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If moEditBox Is Nothing Then
        If KeyCode = vbKeyF2 Then
            KeyCode = 0
            BeginEditInternal False
        End If
    Else
        HandleEditorKeyDown KeyCode, Shift
    End If
End Sub

' Commit the active plain-text or markup edit.
Public Sub CommitEdit()
    Dim sValue As String
    Dim sPreviousMarkup As String
    If moEditBox Is Nothing Then Exit Sub
    sValue = moEditBox.Text
    If mbSourceMode Then
        sPreviousMarkup = ToMarkup
        If Not ParseMarkup(sValue) Then
            ParseMarkup sPreviousMarkup
            Exit Sub
        End If
    Else
        ApplyPlainTextEdit sValue
    End If
    EndEditControls
    Render
    SetFocus
    RaiseEvent Change
    RaiseEvent EditCommitted(mbSourceMode)
End Sub

' Cancel the active edit and restore the last committed display.
Public Sub CancelEdit()
    If moEditBox Is Nothing Then Exit Sub
    EndEditControls
    Render
    SetFocus
    RaiseEvent EditCancelled
End Sub

' Reposition the editor and rendered runs after the host Frame changes size.
Public Sub Resize()
    If Not moEditBox Is Nothing Then
        moEditBox.Left = 2
        moEditBox.Top = 2
        moEditBox.Width = MaxSingle(1, moHost.InsideWidth - 4)
        moEditBox.Height = MaxSingle(1, moHost.InsideHeight - 4)
    End If
    Render
End Sub

' Remove generated controls and release host references.
Public Sub Terminate()
    EndEditControls
    RemoveControls
    HideFocusBorder
    If Not moFocusProxy Is Nothing Then
        On Error Resume Next
        moHost.Controls.Remove moFocusProxy.Name
        On Error GoTo 0
    End If
    Set moFocusProxy = Nothing
    Set moHost = Nothing
End Sub

' Create the hidden TextBox that receives host keyboard input.
Private Sub EnsureFocusProxy()
    If moFocusProxy Is Nothing And Not moHost Is Nothing Then
        Set moFocusProxy = moHost.Controls.Add("Forms.TextBox.1", ControlName("Focus"), False)
        With moFocusProxy
            .Left = 0
            .Top = 0
            .Width = 0.1
            .Height = 0.1
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            .TabStop = True
            .Visible = True
            .ZOrder 1
        End With
    End If
End Sub

' Create and initialize the visible temporary editor.
Private Sub BeginEditInternal(ByVal bSource As Boolean)
    If mbReadOnly Or moHost Is Nothing Then Exit Sub
    EndEditControls
    ShowFocusBorder
    Set moEditBox = moHost.Controls.Add("Forms.TextBox.1", ControlName("Edit"), False)
    mbSourceMode = bSource
    If bSource Then
        moEditBox.Text = ToMarkup
    Else
        moEditBox.Text = msText
    End If
    With moEditBox
        .Left = 2
        .Top = 2
        .Width = MaxSingle(1, moHost.InsideWidth - 4)
        .Height = MaxSingle(1, moHost.InsideHeight - 4)
        .MultiLine = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsBoth
        .Visible = True
        .ZOrder 0
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    HideRenderedControls
End Sub

' Remove the temporary editor without changing committed content.
Private Sub EndEditControls()
    If moEditBox Is Nothing Then Exit Sub
    On Error Resume Next
    moHost.Controls.Remove moEditBox.Name
    On Error GoTo 0
    Set moEditBox = Nothing
End Sub

' Forward editor key presses to the shared shortcut handler.
Private Sub moEditBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEditorKeyDown KeyCode, Shift
End Sub

' Keep the focus border visible while the editor is active.
Private Sub moEditBox_Enter()
    ShowFocusBorder
End Sub

' Restore the host border when the editor loses focus.
Private Sub moEditBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    HideFocusBorder
End Sub

' Forward proxy key presses so F2 works while the Frame is active.
Private Sub moFocusProxy_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleKeyDown KeyCode, Shift
End Sub

' Show the focus border when the proxy receives focus.
Private Sub moFocusProxy_Enter()
    ShowFocusBorder
End Sub

' Restore the host border when the proxy loses focus.
Private Sub moFocusProxy_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    HideFocusBorder
End Sub

' Implement Escape, Ctrl+Enter, F2, and the markup toggle shortcuts.
Private Sub HandleEditorKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        CancelEdit
    ElseIf KeyCode = vbKeyReturn And (Shift And MC_KEY_CTRL) <> 0 Then
        KeyCode = 0
        CommitEdit
    ElseIf KeyCode = vbKeyF2 Then
        KeyCode = 0
        ToggleEditMode
    ElseIf (Shift And MC_KEY_CTRL) <> 0 And (Shift And MC_KEY_SHIFT) <> 0 And _
           (KeyCode = 188 Or KeyCode = 191) Then
        KeyCode = 0
        ToggleEditMode
    End If
End Sub

' Apply the visual focus border used by the host UserForm.
Private Sub ShowFocusBorder()
    If moHost Is Nothing Then Exit Sub
    moHost.BorderStyle = fmBorderStyleSingle
    moHost.BorderColor = RGB(31, 95, 146)
End Sub

' Restore the border settings captured from the host Frame.
Private Sub HideFocusBorder()
    If moHost Is Nothing Then Exit Sub
    moHost.BorderStyle = mlHostBorderStyle
    moHost.BorderColor = mlHostBorderColor
End Sub

' Replace text and initialize one format slot per character.
Private Sub ResetText(ByVal sValue As String)
    Dim lPosition As Long
    msText = sValue
    mlTextLength = Len(sValue)
    If mlTextLength = 0 Then
        Erase mvFormats
        Exit Sub
    End If
    ReDim mvFormats(1 To mlTextLength)
    For lPosition = 1 To mlTextLength
        mvFormats(lPosition) = Empty
    Next lPosition
End Sub

' Merge a plain-text edit while preserving unchanged character formatting.
Private Sub ApplyPlainTextEdit(ByVal sValue As String)
    Dim sOldText As String
    Dim lOldLength As Long
    Dim lNewLength As Long
    Dim lPrefix As Long
    Dim lSuffix As Long
    Dim lPosition As Long
    Dim lOldPosition As Long
    Dim vFormat As Variant
    Dim vInsertedFormat As Variant
    Dim vOldFormats() As Variant

    sOldText = msText
    lOldLength = mlTextLength
    lNewLength = Len(sValue)
    Do While lPrefix < lOldLength And lPrefix < lNewLength
        If Mid$(sOldText, lPrefix + 1, 1) <> Mid$(sValue, lPrefix + 1, 1) Then Exit Do
        lPrefix = lPrefix + 1
    Loop
    Do While lSuffix < lOldLength - lPrefix And lSuffix < lNewLength - lPrefix
        If Mid$(sOldText, lOldLength - lSuffix, 1) <> Mid$(sValue, lNewLength - lSuffix, 1) Then Exit Do
        lSuffix = lSuffix + 1
    Loop

    If lOldLength > 0 Then
        ReDim vOldFormats(1 To lOldLength)
        For lPosition = 1 To lOldLength
            vOldFormats(lPosition) = CloneFormat(CharacterFormat(lPosition))
        Next lPosition
    End If

    If lNewLength = 0 Then
        msText = vbNullString
        mlTextLength = 0
        Erase mvFormats
        Exit Sub
    End If

    If lPrefix > 0 Then
        vInsertedFormat = vOldFormats(lPrefix)
    ElseIf lSuffix > 0 Then
        vInsertedFormat = vOldFormats(lOldLength - lSuffix + 1)
    Else
        vInsertedFormat = DefaultFormat()
    End If

    msText = sValue
    mlTextLength = lNewLength
    ReDim mvFormats(1 To lNewLength)
    For lPosition = 1 To lNewLength
        If lPosition <= lPrefix Then
            lOldPosition = lPosition
            vFormat = vOldFormats(lOldPosition)
        ElseIf lPosition > lNewLength - lSuffix Then
            lOldPosition = lOldLength - lSuffix + (lPosition - (lNewLength - lSuffix))
            vFormat = vOldFormats(lOldPosition)
        Else
            vFormat = vInsertedFormat
        End If
        mvFormats(lPosition) = CloneFormat(vFormat)
    Next lPosition
End Sub

' Append one parsed character and clone its active format state.
Private Sub AppendCharacter(ByVal sCharacter As String, ByVal vFormat As Variant)
    mlTextLength = mlTextLength + 1
    msText = msText & sCharacter
    If mlTextLength = 1 Then
        ReDim mvFormats(1 To 1)
    Else
        ReDim Preserve mvFormats(1 To mlTextLength)
    End If
    mvFormats(mlTextLength) = CloneFormat(vFormat)
End Sub

' Parse escaped literals, tags, and text into the internal model.
Private Function ParseMarkup(ByVal sMarkup As String) As Boolean
    Dim lPosition As Long
    Dim lTagEnd As Long
    Dim sTag As String
    Dim sCharacter As String
    Dim vFormat As Variant
    Dim oFormatStack As Collection
    Dim sError As String

    Set oFormatStack = New Collection
    ResetText vbNullString
    vFormat = DefaultFormat()
    lPosition = 1
    Do While lPosition <= Len(sMarkup)
        If Mid$(sMarkup, lPosition, 1) = "<" Then
            lTagEnd = InStr(lPosition + 1, sMarkup, ">")
            If lTagEnd = 0 Then
                sError = "A markup tag is missing its closing >."
                GoTo ParseFailed
            End If
            sTag = Mid$(sMarkup, lPosition + 1, lTagEnd - lPosition - 1)
            If Not ApplyTag(sTag, vFormat, oFormatStack, sError) Then GoTo ParseFailed
            lPosition = lTagEnd + 1
        ElseIf Mid$(sMarkup, lPosition, 1) = "&" And LCase$(Mid$(sMarkup, lPosition, 4)) = "&lt;" Then
            AppendCharacter "<", vFormat
            lPosition = lPosition + 4
        ElseIf Mid$(sMarkup, lPosition, 1) = "&" And LCase$(Mid$(sMarkup, lPosition, 4)) = "&gt;" Then
            AppendCharacter ">", vFormat
            lPosition = lPosition + 4
        ElseIf Mid$(sMarkup, lPosition, 1) = "&" And LCase$(Mid$(sMarkup, lPosition, 5)) = "&amp;" Then
            AppendCharacter "&", vFormat
            lPosition = lPosition + 5
        Else
            sCharacter = Mid$(sMarkup, lPosition, 1)
            AppendCharacter sCharacter, vFormat
            lPosition = lPosition + 1
        End If
    Loop
    If oFormatStack.Count > 0 Then
        sError = "A markup tag is missing its closing tag."
        GoTo ParseFailed
    End If
    ParseMarkup = True
    Exit Function

ParseFailed:
    ResetText vbNullString
    RaiseEvent ParseError(sError)
End Function

' Apply one opening or closing tag to the active format state.
Private Function ApplyTag(ByVal sRawTag As String, ByRef vFormat As Variant, ByVal oStack As Collection, ByRef sError As String) As Boolean
    Dim sTag As String
    Dim sValue As String
    Dim vPrevious As Variant
    sTag = LCase$(Trim$(sRawTag))
    If Left$(sTag, 1) = "/" Then
        If oStack.Count = 0 Then sError = "Unexpected closing tag: " & sRawTag: Exit Function
        vFormat = CloneFormat(oStack(oStack.Count))
        oStack.Remove oStack.Count
        ApplyTag = True
        Exit Function
    End If
    vPrevious = vFormat
    If sTag = "b" Or sTag = "strong" Then
        vFormat = FormatWithValue(vFormat, MC_BOLD, True)
    ElseIf sTag = "i" Or sTag = "em" Then
        vFormat = FormatWithValue(vFormat, MC_ITALIC, True)
    ElseIf sTag = "u" Then
        vFormat = FormatWithValue(vFormat, MC_UNDERLINE, True)
    ElseIf sTag = "s" Or sTag = "strike" Then
        vFormat = FormatWithValue(vFormat, MC_STRIKETHROUGH, True)
    ElseIf Left$(sTag, 6) = "size=" Then
        vFormat = FormatWithValue(vFormat, MC_FONT_SIZE, CSng(Mid$(sTag, 6)))
    ElseIf Left$(sTag, 5) = "font=" Then
        vFormat = FormatWithValue(vFormat, MC_FONT_NAME, Mid$(sTag, 6))
    ElseIf Left$(sTag, 11) = "font color=" Then
        vFormat = FormatWithValue(vFormat, MC_FORE_COLOR, ParseColor(Mid$(sTag, 12)))
    ElseIf Left$(sTag, 5) = "span " Then
        sValue = StyleValue(sTag)
        vFormat = ApplyStyleText(vFormat, sValue)
    Else
        sError = "Unsupported markup tag: <" & sRawTag & ">"
        Exit Function
    End If
    oStack.Add CloneFormat(vPrevious)
    ApplyTag = True
End Function

' Apply CSS-like style declarations from a span tag.
Private Function ApplyStyleText(ByVal vFormat As Variant, ByVal sStyle As String) As Variant
    Dim vResult As Variant
    vResult = vFormat
    If InStr(1, sStyle, "font-size:", vbTextCompare) > 0 Then vResult = FormatWithValue(vResult, MC_FONT_SIZE, CSng(ReadStyleValue(sStyle, "font-size:")))
    If InStr(1, sStyle, "color:", vbTextCompare) > 0 Then vResult = FormatWithValue(vResult, MC_FORE_COLOR, ParseColor(ReadStyleValue(sStyle, "color:")))
    If InStr(1, sStyle, "background-color:", vbTextCompare) > 0 Then vResult = FormatWithValue(vResult, MC_BACK_COLOR, ParseColor(ReadStyleValue(sStyle, "background-color:")))
    ApplyStyleText = vResult
End Function

' Extract the style attribute from a span tag.
Private Function StyleValue(ByVal sTag As String) As String
    Dim lStart As Long
    lStart = InStr(1, sTag, "style=", vbTextCompare)
    If lStart = 0 Then StyleValue = vbNullString Else StyleValue = Replace$(Mid$(sTag, lStart + 6), Chr$(34), vbNullString)
End Function

' Read one semicolon-delimited style value safely.
Private Function ReadStyleValue(ByVal sStyle As String, ByVal sKey As String) As String
    Dim lStart As Long
    Dim lEnd As Long
    lStart = 1
    Do
        lStart = InStr(lStart, sStyle, sKey, vbTextCompare)
        If lStart = 0 Then Exit Function
        If lStart = 1 Then
            Exit Do
        ElseIf Mid$(sStyle, lStart - 1, 1) = ";" Then
            Exit Do
        End If
        lStart = lStart + Len(sKey)
    Loop
    lStart = lStart + Len(sKey)
    lEnd = InStr(lStart, sStyle, ";")
    If lEnd = 0 Then lEnd = Len(sStyle) + 1
    ReadStyleValue = Trim$(Mid$(sStyle, lStart, lEnd - lStart))
End Function

' Convert #RRGGBB markup into an MSForms color value.
Private Function ParseColor(ByVal sValue As String) As Long
    Dim sHex As String
    sHex = Replace$(Replace$(Replace$(Trim$(sValue), Chr$(34), vbNullString), "#", vbNullString), ";", vbNullString)
    If Len(sHex) <> 6 Then Err.Raise 5, TypeName(Me), "Colors must use #RRGGBB."
    ParseColor = RGB(CLng("&H" & Mid$(sHex, 1, 2)), CLng("&H" & Mid$(sHex, 3, 2)), CLng("&H" & Mid$(sHex, 5, 2)))
End Function

' Copy one format property into a complete format array.
Private Function FormatWithValue(ByVal vFormat As Variant, ByVal lProperty As Long, ByVal vValue As Variant) As Variant
    Dim vResult() As Variant
    Dim lIndex As Long

    If lProperty < 0 Or lProperty >= MC_FORMAT_COUNT Then
        Err.Raise 5, TypeName(Me), "Formatting property index is outside the format array."
    End If

    ReDim vResult(0 To MC_FORMAT_COUNT - 1)
    If IsEmpty(vFormat) Then vFormat = DefaultFormat()
    For lIndex = 0 To MC_FORMAT_COUNT - 1
        vResult(lIndex) = vFormat(lIndex)
    Next
    vResult(lProperty) = vValue
    FormatWithValue = vResult
End Function

' Return a complete format state for one character position.
Private Function CharacterFormat(ByVal lPosition As Long) As Variant
    Dim vFormat As Variant
    vFormat = DefaultFormat()

    If lPosition >= 1 And lPosition <= mlTextLength Then
        If Not IsEmpty(mvFormats(lPosition)) Then vFormat = mvFormats(lPosition)
    End If
    CharacterFormat = vFormat
End Function

' Create the default format state used by plain characters and new text.
Private Function DefaultFormat() As Variant
    Dim vFormat() As Variant
    ReDim vFormat(0 To MC_FORMAT_COUNT - 1)
    vFormat(MC_FONT_NAME) = msDefaultFontName
    vFormat(MC_FONT_SIZE) = msDefaultFontSize
    vFormat(MC_BOLD) = False
    vFormat(MC_ITALIC) = False
    vFormat(MC_UNDERLINE) = False
    vFormat(MC_STRIKETHROUGH) = False
    vFormat(MC_WEIGHT) = Empty
    vFormat(MC_CHARSET) = Empty
    vFormat(MC_FORE_COLOR) = mlDefaultForeColor
    vFormat(MC_BACK_COLOR) = mlDefaultBackColor
    DefaultFormat = vFormat
End Function

' Clone a format array before storing it in another owner.
Private Function CloneFormat(ByVal vFormat As Variant) As Variant
    Dim vClone() As Variant
    Dim lIndex As Long
    ReDim vClone(0 To MC_FORMAT_COUNT - 1)
    For lIndex = 0 To MC_FORMAT_COUNT - 1
        vClone(lIndex) = vFormat(lIndex)
    Next lIndex
    CloneFormat = vClone
End Function

' Return one property from a character format state.
Private Function CharacterFormatValue(ByVal lPosition As Long, ByVal lProperty As Long) As Variant
    Dim vFormat As Variant
    vFormat = CharacterFormat(lPosition)
    CharacterFormatValue = vFormat(lProperty)
End Function

' Compare two character format states for run coalescing.
Private Function SameFormat(ByVal lFirst As Long, ByVal lSecond As Long) As Boolean
    Dim lProperty As Long
    Dim vFirst As Variant
    Dim vSecond As Variant
    vFirst = CharacterFormat(lFirst)
    vSecond = CharacterFormat(lSecond)
    For lProperty = 0 To MC_FORMAT_COUNT - 1
        If Not ValuesEqual(vFirst(lProperty), vSecond(lProperty)) Then Exit Function
    Next lProperty
    SameFormat = True
End Function

' Serialize the internal text and formatting into canonical markup.
Private Function ToMarkup() As String
    Dim lPosition As Long
    Dim lEnd As Long
    Dim sResult As String
    Dim vFormat As Variant
    lPosition = 1
    Do While lPosition < mlTextLength + 1
        lEnd = lPosition
        Do While lEnd < mlTextLength
            If Not SameFormat(lPosition, lEnd + 1) Then Exit Do
            lEnd = lEnd + 1
        Loop
        vFormat = CharacterFormat(lPosition)
        sResult = sResult & OpenMarkup(vFormat)
        sResult = sResult & EscapeMarkup(Mid$(msText, lPosition, lEnd - lPosition + 1))
        sResult = sResult & CloseMarkup(vFormat)
        lPosition = lEnd + 1
    Loop
    ToMarkup = sResult
End Function

' Open canonical tags for one contiguous formatting run.
Private Function OpenMarkup(ByVal vFormat As Variant) As String
    Dim sResult As String
    If CBool(vFormat(MC_BOLD)) Then sResult = sResult & "<b>"
    If CBool(vFormat(MC_ITALIC)) Then sResult = sResult & "<i>"
    If CBool(vFormat(MC_UNDERLINE)) Then sResult = sResult & "<u>"
    If CBool(vFormat(MC_STRIKETHROUGH)) Then sResult = sResult & "<s>"
    If CStr(vFormat(MC_FONT_NAME)) <> msDefaultFontName Then sResult = sResult & "<font=" & CStr(vFormat(MC_FONT_NAME)) & ">"
    If CSng(vFormat(MC_FONT_SIZE)) <> msDefaultFontSize Then sResult = sResult & "<size=" & CStr(vFormat(MC_FONT_SIZE)) & ">"
    If CLng(vFormat(MC_FORE_COLOR)) <> mlDefaultForeColor Then sResult = sResult & "<font color=" & Chr$(34) & ColorMarkup(CLng(vFormat(MC_FORE_COLOR))) & Chr$(34) & ">"
    If CLng(vFormat(MC_BACK_COLOR)) <> mlDefaultBackColor Then sResult = sResult & "<span style=" & Chr$(34) & "background-color:" & ColorMarkup(CLng(vFormat(MC_BACK_COLOR))) & Chr$(34) & ">"
    OpenMarkup = sResult
End Function

' Close canonical tags for one contiguous formatting run.
Private Function CloseMarkup(ByVal vFormat As Variant) As String
    Dim sResult As String
    If CLng(vFormat(MC_BACK_COLOR)) <> mlDefaultBackColor Then sResult = sResult & "</span>"
    If CLng(vFormat(MC_FORE_COLOR)) <> mlDefaultForeColor Then sResult = sResult & "</font>"
    If CSng(vFormat(MC_FONT_SIZE)) <> msDefaultFontSize Then sResult = sResult & "</size>"
    If CStr(vFormat(MC_FONT_NAME)) <> msDefaultFontName Then sResult = sResult & "</font>"
    If CBool(vFormat(MC_STRIKETHROUGH)) Then sResult = sResult & "</s>"
    If CBool(vFormat(MC_UNDERLINE)) Then sResult = sResult & "</u>"
    If CBool(vFormat(MC_ITALIC)) Then sResult = sResult & "</i>"
    If CBool(vFormat(MC_BOLD)) Then sResult = sResult & "</b>"
    CloseMarkup = sResult
End Function

' Escape literal markup characters before serialization.
Private Function EscapeMarkup(ByVal sValue As String) As String
    EscapeMarkup = Replace$(Replace$(Replace$(sValue, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

' Convert an MSForms color value into #RRGGBB markup.
Private Function ColorMarkup(ByVal lColor As Long) As String
    ColorMarkup = "#" & Right$("0" & Hex$(lColor And &HFF&), 2) & _
        Right$("0" & Hex$((lColor \ &H100&) And &HFF&), 2) & _
        Right$("0" & Hex$((lColor \ &H10000) And &HFF&), 2)
End Function

' Recreate the visible formatted runs inside the host Frame.
Private Sub Render()
    Dim lPosition As Long
    Dim lEnd As Long
    Dim sRun As String
    Dim sName As String
    Dim oLabel As MSForms.Label
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngLineHeight As Single
    Dim sngDefaultLineHeight As Single
    Dim sngContentWidth As Single
    If moHost Is Nothing Or mbUpdating Then Exit Sub
    RemoveControls
    If mlTextLength = 0 Then
        SetScrollbars 0, 0
        Exit Sub
    End If
    sngLeft = 3
    sngTop = 3
    sngDefaultLineHeight = MaxSingle(12, moHost.Font.Size * 1.25)
    sngLineHeight = sngDefaultLineHeight
    lPosition = 1
    Do While lPosition <= mlTextLength
        If Mid$(msText, lPosition, 1) = vbCr Or Mid$(msText, lPosition, 1) = vbLf Then
            If Mid$(msText, lPosition, 1) = vbCr And lPosition < mlTextLength Then
                If Mid$(msText, lPosition + 1, 1) = vbLf Then lPosition = lPosition + 1
            End If
            sngLeft = 3
            sngTop = sngTop + sngLineHeight
            sngLineHeight = sngDefaultLineHeight
            lPosition = lPosition + 1
        Else
            lEnd = lPosition
            Do While lEnd < mlTextLength
                If Mid$(msText, lEnd + 1, 1) = vbCr Or Mid$(msText, lEnd + 1, 1) = vbLf Then Exit Do
                If Not SameFormat(lPosition, lEnd + 1) Then Exit Do
                lEnd = lEnd + 1
            Loop
            sRun = Mid$(msText, lPosition, lEnd - lPosition + 1)
            sName = ControlName("Run" & CStr(lPosition))
            Set oLabel = moHost.Controls.Add("Forms.Label.1", sName, False)
            TrackControl sName
            With oLabel
                .Caption = sRun
                .WordWrap = False
                .Left = sngLeft
                .Top = sngTop
                ApplyFormat oLabel, lPosition
                .AutoSize = True
                .Visible = True
                .ZOrder 0
            End With
            If oLabel.Height > sngLineHeight Then sngLineHeight = oLabel.Height
            sngLeft = sngLeft + oLabel.Width
            If sngLeft > sngContentWidth Then sngContentWidth = sngLeft
            lPosition = lEnd + 1
        End If
    Loop
    SetScrollbars sngContentWidth + 3, sngTop + sngLineHeight + 3
End Sub

' Size the virtual canvas and enable only the scrollbars required by rendered content.
Private Sub SetScrollbars(ByVal sngContentWidth As Single, ByVal sngContentHeight As Single)
    Dim lPass As Long
    Dim lBars As Long

    With moHost
        .ScrollWidth = sngContentWidth
        .ScrollHeight = sngContentHeight
        For lPass = 1 To 2
            lBars = fmScrollBarsNone
            If .InsideWidth < .ScrollWidth Then lBars = fmScrollBarsHorizontal
            If .InsideHeight < .ScrollHeight Then lBars = lBars + fmScrollBarsVertical
            If .ScrollBars = lBars Then Exit For
            .ScrollBars = lBars
        Next lPass
    End With
End Sub

' Apply a character format state to one dynamic label.
Private Sub ApplyFormat(ByVal oLabel As MSForms.Label, ByVal lPosition As Long)
    Dim vFormat As Variant
    vFormat = CharacterFormat(lPosition)
    With oLabel
        .Font.Name = CStr(vFormat(MC_FONT_NAME))
        .Font.Size = CSng(vFormat(MC_FONT_SIZE))
        If CBool(vFormat(MC_BOLD)) Then .Font.Bold = True Else .Font.Bold = False
        If CBool(vFormat(MC_ITALIC)) Then .Font.Italic = True Else .Font.Italic = False
        If CBool(vFormat(MC_UNDERLINE)) Then .Font.Underline = True Else .Font.Underline = False
        If CBool(vFormat(MC_STRIKETHROUGH)) Then .Font.Strikethrough = True Else .Font.Strikethrough = False
        .ForeColor = CLng(vFormat(MC_FORE_COLOR))
        .BackColor = CLng(vFormat(MC_BACK_COLOR))
        If CLng(vFormat(MC_BACK_COLOR)) = mlDefaultBackColor Then
            .BackStyle = fmBackStyleTransparent
        Else
            .BackStyle = fmBackStyleOpaque
        End If
    End With
End Sub

' Hide rendered labels while the temporary editor is active.
Private Sub HideRenderedControls()
    Dim vName As Variant
    If mcolControlNames Is Nothing Then Exit Sub
    For Each vName In mcolControlNames
        moHost.Controls(CStr(vName)).Visible = False
    Next vName
End Sub

' Remove all dynamic labels tracked by the control.
Private Sub RemoveControls()
    Dim vName As Variant
    If mcolControlNames Is Nothing Then Set mcolControlNames = New Collection
    On Error Resume Next
    For Each vName In mcolControlNames
        moHost.Controls.Remove CStr(vName)
    Next vName
    On Error GoTo 0
    Set mcolControlNames = New Collection
End Sub

' Add a dynamic control name to the cleanup collection.
Private Sub TrackControl(ByVal sName As String)
    If mcolControlNames Is Nothing Then Set mcolControlNames = New Collection
    mcolControlNames.Add sName
End Sub

' Build a stable name for a generated MSForms control.
Private Function ControlName(ByVal sSuffix As String) As String
    ControlName = "rt_" & sSuffix
End Function

' Validate one-based character range arguments.
Private Sub ValidateRange(ByVal lStart As Long, ByVal lLength As Long)
    If lStart < 1 Or lStart > mlTextLength Then Err.Raise 5, TypeName(Me), "Character start is outside the text."
    If lLength = -1 Then Exit Sub
    If lLength < 1 Or lStart + lLength - 1 > mlTextLength Then Err.Raise 5, TypeName(Me), "Character range is outside the text."
End Sub

' Map a public formatting property name to its internal slot.
Private Function FormatPropertyIndex(ByVal sProperty As String) As Long
    Select Case LCase$(sProperty)
        Case "fontname": FormatPropertyIndex = MC_FONT_NAME
        Case "fontsize": FormatPropertyIndex = MC_FONT_SIZE
        Case "bold": FormatPropertyIndex = MC_BOLD
        Case "italic": FormatPropertyIndex = MC_ITALIC
        Case "underline": FormatPropertyIndex = MC_UNDERLINE
        Case "strikethrough": FormatPropertyIndex = MC_STRIKETHROUGH
        Case "weight": FormatPropertyIndex = MC_WEIGHT
        Case "charset": FormatPropertyIndex = MC_CHARSET
        Case "forecolor": FormatPropertyIndex = MC_FORE_COLOR
        Case "backcolor": FormatPropertyIndex = MC_BACK_COLOR
        Case Else: Err.Raise 5, TypeName(Me), "Unknown formatting property: " & sProperty
    End Select
End Function

' Compare Variant format values without treating Empty as text.
Private Function ValuesEqual(ByVal vLeft As Variant, ByVal vRight As Variant) As Boolean
    If IsEmpty(vLeft) And IsEmpty(vRight) Then ValuesEqual = True Else ValuesEqual = (CStr(vLeft) = CStr(vRight))
End Function

' Return the larger of two Single values.
Private Function MaxSingle(ByVal sngLeft As Single, ByVal sngRight As Single) As Single
    If sngLeft > sngRight Then MaxSingle = sngLeft Else MaxSingle = sngRight
End Function

' Return the smaller of two Long values.
Private Function MinLong(ByVal lLeft As Long, ByVal lRight As Long) As Long
    If lLeft < lRight Then MinLong = lLeft Else MinLong = lRight
End Function
