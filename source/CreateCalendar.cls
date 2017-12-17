VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cs As New ConfigureSheet

Sub NewCalendar(currentYear As Integer, currentMonth As Integer, startRow As Long, startCol As Long)
    
    Dim rngCalendarWhole As Range
    Dim rngCalendarHeaders As Range
    Dim rngCalendarBody As Range

    Dim startDayName As String
    Dim startDayOrdinal As Integer
    Dim numberOfDays As Integer

    startDayName = GetNameOfFirstDayOfMonth(currentYear, currentMonth)
    startDayOrdinal = GetOrdinalOfFirstDayOfMonth(currentYear, currentMonth)
    numberOfDays = HowManyDaysInMonth(DateSerial(currentYear, currentMonth, 1))
    
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
