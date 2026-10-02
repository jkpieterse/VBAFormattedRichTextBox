VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRichTextLinkLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Label As MSForms.Label
Attribute Label.VB_VarHelpID = -1

Private mOwner As clsRichTextBox
Private msLinkTarget As String

Public Property Set Owner(ByVal oOwner As clsRichTextBox)
    Set mOwner = oOwner
End Property

Public Property Let LinkTarget(ByVal sLinkTarget As String)
    msLinkTarget = sLinkTarget
End Property

Private Sub Label_Click()
    If Not mOwner Is Nothing Then mOwner.HandleLinkClick msLinkTarget
End Sub