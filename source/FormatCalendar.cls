VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub SetTitle(startRow As Long, startCol As Long, currentMonth As Integer, currentYear As Integer)

    Dim rngTitle As Range
    Set rngTitle = Range(Cells(startRow - 1, startCol), Cells(startRow - 1, startCol))
    
    Dim rngYear As Range
    Set rngYear = rngTitle.Offset(0, 1)
    
    With rngTitle
        .Value = VBA.MonthName(currentMonth, False)
        .Font.Size = 20
    End With
    
    With rngYear
        .Value = currentYear
        .Font.Size = 20
        .EntireRow.Insert
    End With

End Sub

Sub PopulateDays(currentYear As Integer, currentMonth As Integer, rngCalendarBody As Range)
    
    Dim daysInThisMonth As Integer
        daysInThisMonth = HowManyDaysInMonth(DateSerial(currentYear, currentMonth, 1)) + 1
    
    Dim firstDayThisMonth As Integer
        firstDayThisMonth = GetOrdinalOfFirstDayOfMonth(currentYear, currentMonth)
    
    Dim calDay As Integer
        calDay = 1
    
    Dim i As Integer
    i = 1
    
    Dim c As Integer
    Dim r As Integer
    
    For r = 1 To 6
        For c = 1 To 7
            If calDay = daysInThisMonth Then
                Exit For
            End If
            
            If i = 1 Then
                c = firstDayThisMonth - 1
                i = i + 1
            Else
                rngCalendarBody.Cells(r, c).Value = calDay
                calDay = calDay + 1
            End If
        Next
        
        If calDay = daysInThisMonth Then
            Exit For
        End If
            
    Next

End Sub

Sub CalendarFormatting(rngCalendarWhole As Range, rngCalendarHeaders As Range, rngCalendarBody As Range, startRow As Long, startCol As Long)
    
    Set rngCalendarWhole = Range(Cells(startRow, startCol), Cells(startRow + 6, startCol + 6))
    Call FormatCalendarWhole(rngCalendarWhole)
    
    Set rngCalendarHeaders = Range(Cells(startRow, startCol), Cells(startRow, startCol + 6))
    Call FormatCalendarHeaders(rngCalendarHeaders)
    
    Set rngCalendarBody = Range(Cells(startRow + 1, startCol), Cells(startRow + 6, startCol + 6))
    Call FormatCalendarBody(rngCalendarBody)
    
End Sub

Private Sub FormatCalendarWhole(rngCalendarWhole As Range)

    With rngCalendarWhole
        .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbBlue
    End With

End Sub

Private Sub FormatCalendarHeaders(rngCalendarHeaders As Range)

    With rngCalendarHeaders
        .Cells.EntireRow.Font.Color = RGB(255, 255, 255)
        .Cells.EntireRow.Font.Size = 16
        .Cells.EntireRow.HorizontalAlignment = xlCenter
        .Cells.Interior.Color = RGB(142, 169, 219)
    End With

End Sub

Private Sub FormatCalendarBody(rngCalendarBody As Range)

    With rngCalendarBody
        .ColumnWidth = 25
        .RowHeight = 90
        .Cells.Font.Size = 16
        .Cells.Columns.VerticalAlignment = xlTop
        .Cells.Borders.Color = vbBlack
        .Cells.Borders.Weight = xlThin
        .Columns(1).Cells.Interior.Color = RGB(255, 230, 193)
        .Columns(7).Cells.Interior.Color = RGB(255, 230, 193)
        .Name = "CalendarBody"
    End With

End Sub

Private Function HowManyDaysInMonth(Optional theDateToMonthify As Date = 0) As Integer

    If theDateToMonthify = 0 Then
        theDateToMonthify = Date
    End If
    
    HowManyDaysInMonth = DateSerial(year(theDateToMonthify), month(theDateToMonthify) + 1, 1) - DateSerial(year(theDateToMonthify), month(theDateToMonthify), 1)
    
End Function

Private Function GetNameOfFirstDayOfMonth(thisYear As Integer, thisMonth As Integer) As String
    GetNameOfFirstDayOfMonth = WeekdayName(Weekday(DateSerial(thisYear, thisMonth, 1)))
End Function

Private Function GetOrdinalOfFirstDayOfMonth(thisYear As Integer, thisMonth As Integer) As Integer
    GetOrdinalOfFirstDayOfMonth = Weekday(DateSerial(thisYear, thisMonth, 1))
End Function

