VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRichTextCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : clsRichTextCharacters
' Purpose: One-based formatting range wrapper for clsRichTextBox.
'---------------------------------------------------------------------------------------
Option Explicit

Private moRichText As clsRichTextBox
Private mlStart As Long
Private mlLength As Long

' Connect this range wrapper to its owning RichText control.
Friend Sub Initialize(ByVal oRichText As clsRichTextBox, ByVal lStart As Long, ByVal lLength As Long)
    Set moRichText = oRichText
    mlStart = lStart
    mlLength = lLength
End Sub

' Return or assign the font name for the range.
Public Property Get FontName() As Variant
    FontName = moRichText.FormatValue(mlStart, mlLength, "FontName")
End Property

' Apply a font name to the range.
Public Property Let FontName(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "FontName", vValue
End Property

' Return or assign the font size for the range.
Public Property Get FontSize() As Variant
    FontSize = moRichText.FormatValue(mlStart, mlLength, "FontSize")
End Property

' Apply a font size to the range.
Public Property Let FontSize(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "FontSize", vValue
End Property

' Return or assign bold formatting for the range.
Public Property Get Bold() As Variant
    Bold = moRichText.FormatValue(mlStart, mlLength, "Bold")
End Property

' Apply bold formatting to the range.
Public Property Let Bold(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "Bold", vValue
End Property

' Return or assign italic formatting for the range.
Public Property Get Italic() As Variant
    Italic = moRichText.FormatValue(mlStart, mlLength, "Italic")
End Property

' Apply italic formatting to the range.
Public Property Let Italic(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "Italic", vValue
End Property

' Return or assign underline formatting for the range.
Public Property Get Underline() As Variant
    Underline = moRichText.FormatValue(mlStart, mlLength, "Underline")
End Property

' Apply underline formatting to the range.
Public Property Let Underline(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "Underline", vValue
End Property

' Return or assign strikethrough formatting for the range.
Public Property Get Strikethrough() As Variant
    Strikethrough = moRichText.FormatValue(mlStart, mlLength, "Strikethrough")
End Property

' Apply strikethrough formatting to the range.
Public Property Let Strikethrough(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "Strikethrough", vValue
End Property

' Return or assign foreground color for the range.
Public Property Get ForeColor() As Variant
    ForeColor = moRichText.FormatValue(mlStart, mlLength, "ForeColor")
End Property

' Apply a foreground color to the range.
Public Property Let ForeColor(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "ForeColor", vValue
End Property

' Return or assign background color for the range.
Public Property Get BackColor() As Variant
    BackColor = moRichText.FormatValue(mlStart, mlLength, "BackColor")
End Property

' Apply a background color to the range.
Public Property Let BackColor(ByVal vValue As Variant)
    moRichText.SetFormatValue mlStart, mlLength, "BackColor", vValue
End Property

' Remove all custom formatting from the range.
Public Sub Clear()
    moRichText.ClearFormat mlStart, mlLength
End Sub

' Release the owning control when this temporary range is destroyed.
Private Sub Class_Terminate()
    Set moRichText = Nothing
End Sub
