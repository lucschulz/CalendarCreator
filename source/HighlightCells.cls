VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighlightCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub HighlightByDate(dayOfMonth As String)

    Range("CalendarBody").Find(dayOfMonth).Interior.Color = RGB(102, 153, 255)

End Sub
