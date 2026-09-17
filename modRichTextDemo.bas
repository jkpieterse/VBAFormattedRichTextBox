Attribute VB_Name = "modRichTextDemo"
'---------------------------------------------------------------------------------------
' File   : modRichTextDemo
' Purpose: Public entry point for the standalone RichText demonstration form.
'---------------------------------------------------------------------------------------
Option Explicit

Public Const GCSAPPNAME As String = "RichText Demo"
Public Const GCSBUILD As String = "1.0"

' Open the standalone demonstration form.
Public Sub DemoNow()
    Dim ufForm As ufRichTextDemo
    Set ufForm = New ufRichTextDemo
    ufForm.AppName = GCSAPPNAME
    ufForm.Show
    Unload ufForm
    Set ufForm = Nothing
End Sub

