VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCalendarFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function HowManyDaysInMonth(Optional theDateToMonthify As Date = 0) As Integer

    If theDateToMonthify = 0 Then
        theDateToMonthify = Date
    End If
    
    HowManyDaysInMonth = DateSerial(Year(theDateToMonthify), Month(theDateToMonthify) + 1, 1) - DateSerial(Year(theDateToMonthify), Month(theDateToMonthify), 1)
    
End Function

Public Function GetNameOfFirstDayOfMonth(thisYear As Integer, thisMonth As Integer) As String
    GetNameOfFirstDayOfMonth = WeekdayName(Weekday(DateSerial(thisYear, thisMonth, 1)))
End Function

Public Function GetOrdinalOfFirstDayOfMonth(thisYear As Integer, thisMonth As Integer) As Integer
    GetOrdinalOfFirstDayOfMonth = Weekday(DateSerial(thisYear, thisMonth, 1))
End Function

Public Property Get DaysOfTheWeek() As Variant
    DaysOfTheWeek = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
End Property
