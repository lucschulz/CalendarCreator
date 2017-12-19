VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigureSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub CreateNewSheet()

    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    With newSheet
        .Activate
    End With
    
    With ActiveWindow
        .DisplayGridlines = False
        .WindowState = xlMaximized
        .Zoom = 48
    End With

End Sub

Sub CreateNewWorkbook()

    Dim newWorkbook As Workbook
    Set newWorkbook = Application.Workbooks.Add
    
End Sub
