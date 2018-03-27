VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSqlParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get NewVarCharParameter(prmName As String, prmSize As Integer, prmValue As String) As ADODB.Parameter
    Set NewVarCharParameter = VarCharParameter(prmName, prmSize, prmValue)
End Property


Public Property Get NewIntegerParameter(prmName As String, prmValue As Integer) As ADODB.Parameter
    Set NewIntegerProperty = IntegerParameter(prmName, prmValue)
End Property


Public Property Get NewDateParameter(prmName As String, prmValue As Date) As ADODB.Parameter
    Set NewDateParameter = DateParameter(prmName, prmValue)
End Property



Private Function VarCharParameter(prmName As String, prmSize As Integer, prmValue As String) As ADODB.Parameter
    Dim prm As New ADODB.Parameter
    With prm
        .Name = prmName
        .Direction = adParamInput
        .Size = prmSize
        .Type = adVarChar
        .Value = prmValue
    End With
    
    Set VarCharParameter = prm
End Function



Private Function IntegerParameter(prmName As String, prmValue As Integer) As ADODB.Parameter
    Dim prm As New ADODB.Parameter
    With prm
        .Name = prmName
        .Direction = adParamInput
        .Type = adInteger
        .Value = prmValue
    End With
    
    Set IntegerParameter = prm
End Function



Private Function DateParameter(prmName As String, prmValue As Date) As ADODB.Parameter
    Dim prm As New ADODB.Parameter
    With prm
        .Name = prmName
        .Direction = adParamInput
        .Type = adDate
        .Value = prmValue
    End With
    
    Set DateParameter = prm
End Function
