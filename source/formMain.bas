VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMain 
   Caption         =   "Calendar Creator"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "formMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub BtnGenerateCalendar_Click()

    Dim currentYear As Integer
    currentYear = CInt(tbYear.Text)
    modEntireYear.GenerateNewCalendar currentYear
    
    'Unload Me
    
'    If cbListOfMonths.ListIndex = -1 Then
'        MsgBox "You must select a month in order to proceed."
'        Exit Sub
'    End If
'
'    Dim currentYear As Integer
'    Dim currentMonth As Integer
'
'    currentYear = CInt(tbYear.Text)
'    currentMonth = cbListOfMonths.ListIndex + 1
'
'    Dim cs As New ConfigureSheet
'    Dim cc As New CreateCalendar
'
'    If checkUseNewWorkbook.Value = True Then
'        Call cs.CreateNewWorkbook
'        Call cc.NewCalendar(currentYear, currentMonth, 2, 2)
'    Else
'        Call cs.CreateNewSheet
'        Call cc.NewCalendar(currentYear, currentMonth, 2, 2)
'    End If
        
End Sub

Private Sub BtnHighlightDates_Click()
    Dim h As New HighlightCells
    Call h.HighlightByDate
End Sub

Private Sub checkEntireYear_Click()
    
    checkEntireYear.Value = False
    MsgBox "Option not yet available."
    Exit Sub
        
    If checkEntireYear.Value = True Then
        cbListOfMonths.Enabled = False
    ElseIf checkEntireYear.Value = False Then
        cbListOfMonths.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    
    PopulateListOfMonths
    
End Sub

Private Sub PopulateListOfMonths()
    
    Dim arrMonths(1 To 12) As String
        
    Dim i As Integer
    For i = 1 To 12
        arrMonths(i) = VBA.MonthName(i)
    Next
    
    cbListOfMonths.List = arrMonths
    
End Sub
