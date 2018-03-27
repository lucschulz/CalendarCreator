VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPopulateComboBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ListOfMonths() As Variant
    ListOfMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
End Property

Public Sub PopulateYears(cb As ComboBox)
    
    Dim listOfYears(0 To 150) As Integer
    Dim i As Integer
    
    For i = 0 To 150
        listOfYears(i) = i + 1920
    Next i
    
    cb.List = listOfYears
    
End Sub

