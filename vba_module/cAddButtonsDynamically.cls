VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAddButtonsDynamically"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AddButtonToFrame(frm As Frame, leftOffset As Integer, topOffset As Integer, buttonName As String, Optional buttonCaption As Variant)
    
    Dim btn As New Control
    Set btn = frm.Controls.Add("Forms.CommandButton.1", buttonName)
        
    With btn
        
        .Height = 20
        .Width = 20
        .left = leftOffset
        .top = topOffset
        
        If IsMissing(buttonCaption) = False Then
            .Caption = buttonCaption
        End If
    
    End With
    
End Sub
