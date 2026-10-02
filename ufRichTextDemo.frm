VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRichTextDemo 
   Caption         =   "RichText Demo 1.0.1"
   ClientHeight    =   9636.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11988
   OleObjectBlob   =   "ufRichTextDemo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRichTextDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : ufRichTextDemo
' Author : Jan Karel Pieterse
' (c)    : Copyright JKP Application Development Services, all rights reserved
' Date   : 16-Sep-26
' Purpose: Demonstrates how to implement the All-VBA Rich Text Box into your project
'---------------------------------------------------------------------------------------
Option Explicit

'=======================================================================================
' MANDATORY HOST IMPLEMENTATION
' Copy this declaration and the routines in this section to the UserForm that owns
' your RichText Frame. Replace frRichText with the name of your own MSForms.Frame.
'=======================================================================================
Private WithEvents mcRichText As clsRichTextBox
Attribute mcRichText.VB_VarHelpID = -1
Private msAppName As String

'=======================================================================================
' MANDATORY HOST ROUTINES: keep this complete block together in the owning UserForm.
' Replace frRichText with the name of the design-time MSForms.Frame on your form.
'=======================================================================================
' Create the control, assign its host Frame, and load initial content.
' The host Frame must already exist on the UserForm at design time.
'=======================================================================================
Private Sub UserForm_Initialize()
    Set mcRichText = New clsRichTextBox
    Set mcRichText.HostFrame = Me.frRichText
    mcRichText.WordWrap = True
    ' Optional demo content: replace or remove this assignment on your own form.
    mcRichText.Markup = SampleMarkup()
End Sub

'=======================================================================================
' MANDATORY: return focus to the RichText control when the host form activates.
' This makes F2 available through the control's hidden keyboard focus proxy.
'=======================================================================================
Private Sub UserForm_Activate()
    If Not mcRichText Is Nothing Then mcRichText.SetFocus
    Me.Caption = Sheet1.Range("A1").Value
End Sub

'=======================================================================================
' MANDATORY: forward host Frame focus transitions to the RichText control.
' This drives the focus border and mirrors the clsTreeView EnterExit pattern.
'=======================================================================================
Private Sub frRichText_Enter()
    If Not mcRichText Is Nothing Then mcRichText.EnterExit False
End Sub

'=======================================================================================
' MANDATORY: tell the control when focus leaves its host Frame.
'=======================================================================================
Private Sub frRichText_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mcRichText Is Nothing Then mcRichText.EnterExit True
End Sub

'=======================================================================================
' MANDATORY: resize the dynamic RichText controls with their host Frame.
'=======================================================================================
Private Sub frRichText_Resize()
    If Not mcRichText Is Nothing Then mcRichText.Resize
End Sub

'=======================================================================================
' Release generated controls before the UserForm unloads.
'=======================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mcRichText Is Nothing Then mcRichText.Terminate
    Set mcRichText = Nothing
End Sub

'=======================================================================================
' END MANDATORY HOST ROUTINES
'=======================================================================================

'---------------------------------------------------------------------------------------
' Optional demo property: host-facing application title.
'---------------------------------------------------------------------------------------
Public Property Let AppName(ByVal sValue As String)
    msAppName = sValue
End Property

'---------------------------------------------------------------------------------------
' Optional host event: react to model changes when dependent UI needs refreshing.
'---------------------------------------------------------------------------------------
Private Sub mcRichText_Change()
    UpdateInfo
End Sub

'---------------------------------------------------------------------------------------
' Optional host event: report invalid markup when source editing is exposed.
' The control keeps the last valid content when parsing fails.
'---------------------------------------------------------------------------------------
Private Sub mcRichText_ParseError(ByVal Message As String)
    Me.labInfo.Caption = "Markup error: " & Message
End Sub

' Route demo links by their target prefix: url: opens a website; sheet: activates a worksheet.
Private Sub mcRichText_LinkClick(ByVal LinkTarget As String)
    Dim lSeparator As Long
    Dim sLinkType As String
    Dim sLinkValue As String
    Dim sWorksheetName As String
    Dim sRangeAddress As String
    Dim oWorksheet As Worksheet
    Dim oRange As Range

    lSeparator = InStr(1, LinkTarget, ":", vbBinaryCompare)
    If lSeparator < 2 Then
        Me.labInfo.Caption = "Unsupported link target: " & LinkTarget
        Exit Sub
    End If
    sLinkType = LCase$(Left$(LinkTarget, lSeparator - 1))
    sLinkValue = Mid$(LinkTarget, lSeparator + 1)
    Select Case sLinkType
        Case "url"
            If Len(sLinkValue) = 0 Then Exit Sub
            ThisWorkbook.FollowHyperlink Address:=sLinkValue
        Case "sheet"
            lSeparator = InStrRev(sLinkValue, "!")
            If lSeparator < 2 Or lSeparator >= Len(sLinkValue) Then
                Me.labInfo.Caption = "Worksheet links must use sheet:<worksheet name>!<range address>."
                Exit Sub
            End If
            sWorksheetName = Trim$(Left$(sLinkValue, lSeparator - 1))
            If Left$(sWorksheetName, 1) = "'" And Right$(sWorksheetName, 1) = "'" Then
                sWorksheetName = Replace$(Mid$(sWorksheetName, 2, Len(sWorksheetName) - 2), "''", "'")
            End If
            sRangeAddress = Trim$(Mid$(sLinkValue, lSeparator + 1))
            On Error Resume Next
            Set oWorksheet = ThisWorkbook.Worksheets(sWorksheetName)
            On Error GoTo 0
            If oWorksheet Is Nothing Then
                Me.labInfo.Caption = "Worksheet not found: " & sWorksheetName
                Exit Sub
            End If
            On Error Resume Next
            Set oRange = oWorksheet.Range(sRangeAddress)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Me.labInfo.Caption = "Invalid worksheet range: " & sLinkValue
                Exit Sub
            End If
            On Error GoTo 0
            oWorksheet.Activate
            oRange.Select
        Case Else
            Me.labInfo.Caption = "Unsupported link type: " & sLinkType
    End Select
End Sub

'---------------------------------------------------------------------------------------
' Optional demo content. Replace this routine with the markup your own form needs.
'---------------------------------------------------------------------------------------
Private Function SampleMarkup() As String
    SampleMarkup = "<size=14><b>This Rich Text Box uses HTML-like markup tags to set character formatting.</b></size>" & vbNewLine & vbNewLine
    SampleMarkup = SampleMarkup & "Use url: targets for websites and sheet:SheetName!A1,A3 for workbook ranges. Quote worksheet names with spaces; comma-separated areas are supported." & vbNewLine
    SampleMarkup = SampleMarkup & "Visit <link=url:https://jkp-ads.com><u><font color=""#0563C1"">jkp-ads.com</font></u></link> or select <link=sheet:Overview!A1,A3><u><font color=""#0563C1"">Overview cells A1 and A3</font></u></link>." & vbNewLine & vbNewLine
    SampleMarkup = SampleMarkup & "Examples:" & vbNewLine & vbNewLine
    SampleMarkup = SampleMarkup & "&lt;b&gt;bold&lt;/b&gt; gives <b>bold</b> text." & vbNewLine
    SampleMarkup = SampleMarkup & "&lt;i&gt;italic&lt;/i&gt; gives <i>italic</i> text." & vbNewLine
    SampleMarkup = SampleMarkup & "&lt;u&gt;underline;/u&gt; gives <u>underlined</u> text." & vbNewLine
    SampleMarkup = SampleMarkup & "&lt;span style=""color:#FF0000""&gt;red&lt;/span&gt; gives <span style=""color:#FF0000"">red</span> text" & vbNewLine
    SampleMarkup = SampleMarkup & "Or, simpler:" & vbNewLine
    SampleMarkup = SampleMarkup & "&lt;font color=""#FF0000""&gt;red&lt;/font&gt; gives <font color=""#FF0000"">red</font> text"
    SampleMarkup = SampleMarkup & vbNewLine & vbNewLine & "<b>Short-cut keys:</b>"
    SampleMarkup = SampleMarkup & vbNewLine & vbNewLine & "F2: Start editing text"
    SampleMarkup = SampleMarkup & vbNewLine & "Ctrl+&lt; or Ctrl+? : Toggle markup/text editing mode"
    SampleMarkup = SampleMarkup & vbNewLine & "Ctrl+Enter : Commit edited text/markup"
    SampleMarkup = SampleMarkup & vbNewLine & "Escape : Cancel editing text/markup"
End Function

'---------------------------------------------------------------------------------------
' Optional demo link; this is not required when embedding the control.
'---------------------------------------------------------------------------------------
Private Sub lblLink_Click()
    ActiveWorkbook.FollowHyperlink "https://jkp-ads.com/articles/vba-rich-text-control.aspx"
End Sub

'---------------------------------------------------------------------------------------
' Optional demo close button; this is not required when embedding the control.
'---------------------------------------------------------------------------------------
Private Sub cmbClose_Click()
    Me.Hide
End Sub

'---------------------------------------------------------------------------------------
' Optional demo command: start plain-text editing.
'---------------------------------------------------------------------------------------
Private Sub cmdEdit_Click()
    mcRichText.BeginEdit
End Sub

'---------------------------------------------------------------------------------------
' Optional demo command: start source/markup editing.
'---------------------------------------------------------------------------------------
Private Sub cmdSource_Click()
    mcRichText.BeginSourceEdit
End Sub

'---------------------------------------------------------------------------------------
' Optional demo command: commit the current edit.
'---------------------------------------------------------------------------------------
Private Sub cmdCommit_Click()
    mcRichText.CommitEdit
    UpdateInfo
End Sub

'---------------------------------------------------------------------------------------
' Optional demo command: cancel the current edit.
'---------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    mcRichText.CancelEdit
    UpdateInfo
End Sub

'---------------------------------------------------------------------------------------
' Optional demo command: restore the sample markup.
'---------------------------------------------------------------------------------------
Private Sub cmdReset_Click()
    mcRichText.Markup = SampleMarkup()
    UpdateInfo
End Sub

'---------------------------------------------------------------------------------------
' Optional demo event: refresh the status panel after a successful edit.
'---------------------------------------------------------------------------------------
Private Sub mcRichText_EditCommitted(ByVal SourceMode As Boolean)
    UpdateInfo
End Sub

'---------------------------------------------------------------------------------------
' Optional demo status display; replace with the host UI or remove it.
'---------------------------------------------------------------------------------------
Private Sub UpdateInfo()
    If mcRichText Is Nothing Then Exit Sub
    Me.labInfo.Caption = "Text length: " & Len(mcRichText.Text) & vbCrLf & _
        "Markup: " & mcRichText.Markup
End Sub

