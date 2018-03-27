VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTagCalendarButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TagCalendarButtons(f As Frame)
    
    Dim btn As Variant
    Dim i As Integer
    i = 0
    
    For Each btn In f.Controls
        If TypeName(btn) = "CommandButton" Then
            btn.Tag = i
            btn.Caption = i
            i = i + 1
        End If
    Next btn
    
'    For Each btn In Me.Controls
'        If btn.Tag = 4 Then
'            btn.Caption = "A"
'        End If
'    Next btn
    
End Sub
