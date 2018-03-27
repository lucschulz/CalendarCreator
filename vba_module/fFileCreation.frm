VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fFileCreation 
   Caption         =   "File Creation"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2805
   OleObjectBlob   =   "fFileCreation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fFileCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    
    Dim tags As New cTagCalendarButtons
    tags.TagCalendarButtons frameDaysOfTheMonth
    
    Debug.Print Me.frameDaysOfTheMonth.Controls(1).Tag
    
    Dim cbs As New cPopulateComboBoxes
    cbMonths.List = cbs.ListOfMonths
    cbs.PopulateYears cbYears
    
End Sub



