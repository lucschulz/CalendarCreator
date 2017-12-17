Attribute VB_Name = "modMain"
Option Explicit

Sub OutputDaysOfTheWeek(currentYear As Integer, currentMonth As Integer)
    
    ClearArea
    
    Dim rngCalendarWhole As Range
    Dim rngCalendarHeaders As Range
    Dim rngCalendarBody As Range

    Dim startDayName As String
    Dim startDayOrdinal As Integer
    Dim numberOfDays As Integer

    startDayName = GetNameOfFirstDayOfMonth(currentYear, currentMonth)
    startDayOrdinal = GetOrdinalOfFirstDayOfMonth(currentYear, currentMonth)
    numberOfDays = HowManyDaysInMonth(DateSerial(currentYear, currentMonth, 1))
    
    Dim startRow As Long
    Dim startCol As Long
    
    startRow = 2
    startCol = 2

    Dim rng As Range
    Set rng = Range(Cells(startRow, startCol), Cells(startRow, startCol + 6))
    
    Dim i As Integer
    For i = 1 To 7
        rng.Cells(startRow - 1, i).Value = WeekdayName(i)
    Next
    
    Dim rngFirstDayOfMonth As Range
    Set rngFirstDayOfMonth = rng.Offset(startRow - 1, startDayOrdinal - 1).Cells(1, 1)
    rngFirstDayOfMonth.Select
    
    Dim formatCal As New FormatCalendar
    Call formatCal.CalendarFormatting(rngCalendarWhole, rngCalendarHeaders, rngCalendarBody, startRow, startCol)
    Call formatCal.PopulateDays(currentYear, currentMonth, rngCalendarBody)
    Call formatCal.SetTitle(startRow, startCol, currentMonth, currentYear)

End Sub

Private Sub ClearArea()
    Dim rngClearAll As Range
    Set rngClearAll = Range(Cells(1, 1), Cells(10, 10))
    
    rngClearAll.Clear
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
