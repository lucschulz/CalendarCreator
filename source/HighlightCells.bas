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

Sub HightlightAbsences(monthToHighlight As String, dayOfMonth As String)

    Range(monthToHighlight).Find(dayOfMonth).Interior.Color = RGB(102, 153, 255)

End Sub


Sub HighlightByDate()

    Range("October").Find("3").Interior.Color = RGB(102, 153, 255)
    Range("October").Find("4").Interior.Color = RGB(102, 153, 255)
    Range("November").Find("12").Interior.Color = RGB(102, 153, 255)
    Range("December").Find("5").Interior.Color = RGB(102, 153, 255)
    Range("January").Find("8").Interior.Color = RGB(102, 153, 255)
    Range("April").Find("11").Interior.Color = RGB(102, 153, 255)
    Range("April").Find("12").Interior.Color = RGB(102, 153, 255)
    Range("April").Find("13").Interior.Color = RGB(102, 153, 255)
    Range("May").Find("7").Interior.Color = RGB(102, 153, 255)
    Range("July").Find("30").Interior.Color = RGB(102, 153, 255)
    
    Range("June").Find("7").Interior.Color = RGB(102, 153, 0)
    Range("June").Find("8").Interior.Color = RGB(102, 153, 0)
    Range("September").Find("10").Interior.Color = RGB(102, 153, 0)
    Range("March").Find("26").Interior.Color = RGB(102, 153, 0)
    Range("August").Find("7").Interior.Color = RGB(102, 153, 0)
    

End Sub

