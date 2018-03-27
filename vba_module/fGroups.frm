VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fGroups 
   Caption         =   "Groups"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3840
   OleObjectBlob   =   "fGroups.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim evts(1 To 35) As New cDynamicEvents

Private Property Get SelectedMonth() As Date
    SelectedMonth = VBA.Date
End Property

Private Sub UserForm_Initialize()
    
    PopulateComboBoxes
    
    Dim addButton As New cAddButtonsDynamically
    Dim cf As New cCalendarFunctions
    
    Dim leftOffset As Integer, topOffset As Integer
    leftOffset = 0
    topOffset = 0
    
    Dim b As Integer
    Dim r As Integer
    
    For r = 0 To 5
        For b = 0 To 6
            addButton.AddButtonToFrame frameCalendar, leftOffset, topOffset, "btn" & cf.DaysOfTheWeek(b) & r
            leftOffset = leftOffset + 20
        Next
        leftOffset = 0
        topOffset = topOffset + 20
    Next
    
    Dim startingDay As String
    startingDay = "btn" & cf.GetNameOfFirstDayOfMonth(2018, 3) & 0
    Me.Controls(startingDay).Caption = 1
    
    Dim row As Integer
    row = 0
    Dim col As Integer
    col = 5
    
    Dim i As Integer
    For i = 0 To cf.HowManyDaysInMonth(VBA.Date) - 2
        Dim nextDay As String
        nextDay = "btn" & cf.DaysOfTheWeek(col) & row
        Me.Controls(nextDay).Caption = i + 2
        Me.Controls(nextDay).Tag = i + 2
        col = col + 1
        
        Set evts(i + 2).HandleButtonClicks = Me.Controls(nextDay)
        
        If col = 7 Then
            row = row + 1
            col = 0
        End If
    Next
    
End Sub

Private Sub PopulateComboBoxes()
    
    Dim cbs As New cPopulateComboBoxes
    cbMonths.List = cbs.ListOfMonths
    cbs.PopulateYears cbYears
    
End Sub








