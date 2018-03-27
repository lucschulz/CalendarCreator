VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDynamicEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents HandleButtonClicks As MSForms.CommandButton
Attribute HandleButtonClicks.VB_VarHelpID = -1

Private Sub HandleButtonClicks_Click()
    
    MsgBox "test"
    HandleButtonClicks.Caption = "A"
    
End Sub
