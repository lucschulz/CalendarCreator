Attribute VB_Name = "modHighlights"
Option Explicit

Sub GetEmployeeAbsences(employeeID As Integer, theYear As Integer)
    
    Dim sqlQuery As String
        sqlQuery = "SELECT * FROM absences WHERE employeeID = @EmployeeID AND YEAR(absentDate) = @Year;"
    
    Dim cn As ADODB.Connection
        Set cn = New ADODB.Connection
    
    Dim rs As ADODB.recordSet
        Set rs = CreateObject("ADODB.RECORDSET")
        
    Dim prmEmpID As ADODB.Parameter
        Set prmEmpID = New ADODB.Parameter
    
    With prmEmpID
        .Name = "@EmployeeID"
        .Direction = adParamInput
        .Type = adInteger
        .Value = employeeID
    End With
    
    Dim prmYear As ADODB.Parameter
        Set prmYear = New ADODB.Parameter
        
    With prmYear
        .Name = "@Year"
        .Direction = adParamInput
        .Type = adInteger
        .Value = theYear
    End With
    
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    cn.Open GetCString()
    With cmd
        .ActiveConnection = cn
        .CommandText = sqlQuery
        .Parameters.Append prmEmpID
        .Parameters.Append prmYear
    End With
    
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    
    
    Dim hi As New HighlightCells
    
    Do While Not rs.EOF
                
        Dim dt As String
        dt = rs.Fields.Item("absentDate")
        
        Call hi.HightlightAbsences(VBA.MonthName(VBA.Month(dt)), VBA.Day(dt))
        
        rs.MoveNext
    Loop
    
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
End Sub


Sub Test()
    GetEmployeeAbsences 14
End Sub

