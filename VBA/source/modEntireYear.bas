Attribute VB_Name = "modEntireYear"
Option Explicit

Sub GenerateNewCalendar(currentYear As Integer)

    Dim cs As New ConfigureSheet
    Call cs.CreateNewSheet
    
    'FIRST ROW - ADD FEB AND MARCH
    AddMonth currentYear, 1, 2, 2
    AddMonth currentYear, 2, 2, 10
    AddMonth currentYear, 3, 2, 18
    AddMonth currentYear, 4, 2, 26
    
    
    'SECOND ROW - APR, MAY, JUNE
    AddMonth currentYear, 5, 12, 2
    AddMonth currentYear, 6, 12, 10
    AddMonth currentYear, 7, 12, 18
    AddMonth currentYear, 8, 12, 26
    
    
    'THIRD ROW - JULY, AUG, SEPT
    AddMonth currentYear, 9, 22, 2
    AddMonth currentYear, 10, 22, 10
    AddMonth currentYear, 11, 22, 18
    AddMonth currentYear, 12, 22, 26
    
End Sub

Private Sub AddMonth(currentYear As Integer, currentMonth As Integer, startRow As Long, startColumn As Long)
    Dim cc As New CreateCalendar
    Call cc.NewCalendar(currentYear, currentMonth, startRow, startColumn)
End Sub
