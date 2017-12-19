VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CStringBuilder(dbPathFilename As String, dbPassword As String) As String
    Dim dbProvider As String
        dbProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
            CStringBuilder = dbProvider & dbPathFilename & ";;Jet OLEDB:Database Password=" & dbPassword & ";"
End Function

Function FillComboBox(comboBox As Object, sql As String)

    Dim recordSet As Variant
        recordSet = GetRecordSet(sql)
    
    With comboBox
        .Clear
            If IsArray(recordSet) Then
                For Each i In recordSet
                comboBox.AddItem (i)
            Next i
        End If
    End With

End Function

Function FillComboBoxUsingSameRecordSet(comboBox As Object, recordSet As Variant)

    With comboBox
        .Clear
            If IsArray(recordSet) Then
                For Each i In recordSet
                comboBox.AddItem (i)
            Next i
        End If
    End With

End Function

Sub FillListView(listview As Object, sql As String)
    
    Dim records As Variant
    records = GetRecordSet(sql)
    
        With listview
            .ListItems.Clear
            If IsArray(records) Then
                For i = 0 To UBound(records, 2)
                    .ListItems.Add , , records(0, i)
                    .ListItems(i + 1).ListSubItems.Add , , "" & records(1, i)
                Next i
            End If
        End With
End Sub

Sub FillListViewGroups(listview As Object, sql As String)

    Dim rs As Variant
        rs = GetRecords(sql)
    
    With listview
        .ListItems.Clear
        If IsArray(rs) Then
            For i = 0 To UBound(rs, 2)
                .ListItems.Add , , rs(0, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & rs(1, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & rs(2, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & rs(4, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & rs(5, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & rs(7, i)
            Next i
        End If
    End With
End Sub

Sub PullQueueFillListView(listview As Object, sql As String)

    Dim recordSet As Variant
        recordSet = GetRecordSet(sql)

    With listview
        .ListItems.Clear
        If IsArray(recordSet) Then
            For i = 0 To UBound(recordSet, 2)
                .ListItems.Add , , recordSet(0, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & recordSet(3, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & recordSet(26, i)
                .ListItems(i + 1).ListSubItems.Add , , "" & recordSet(30, i)
            Next i
        End If
    End With
End Sub

Sub RunSqlCommand(sql As String)
    Dim cx As ADODB.Connection
        Set cx = New ADODB.Connection
    
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
    
    cx.Open GetCString()
        With cmd
            .ActiveConnection = cx
            .CommandText = sql
            .Execute
        End With
    cx.Close
    
    Set cx = Nothing
    Exit Sub
    
End Sub

Sub RunSqlWithOneParameter(sql As String, prm As ADODB.Parameter)
    Dim cx As ADODB.Connection
        Set cx = New ADODB.Connection
    
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
            cmd.Parameters.Append prm
                cx.Open GetCString()
                With cmd
                    .ActiveConnection = cx
                    .CommandText = sql
                    .Execute
                End With
            cx.Close
        Set cx = Nothing
End Sub

Sub RunSqlWithTwoParameter(sql As String, prm As ADODB.Parameter, prm2 As ADODB.Parameter)
    Dim cx As ADODB.Connection
        Set cx = New ADODB.Connection
    
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
            cmd.Parameters.Append prm
            cmd.Parameters.Append prm2
                cx.Open GetCString()
                With cmd
                    .ActiveConnection = cx
                    .CommandText = sql
                    .Execute
                End With
            cx.Close
        Set cx = Nothing
End Sub

Sub RunSqlWithThreeParameter(sql As String, prm As ADODB.Parameter, prm2 As ADODB.Parameter, prm3 As ADODB.Parameter)
    Dim cx As ADODB.Connection
        Set cx = New ADODB.Connection
    
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
            cmd.Parameters.Append prm
            cmd.Parameters.Append prm2
            cmd.Parameters.Append prm3
                cx.Open GetCString()
                With cmd
                    .ActiveConnection = cx
                    .CommandText = sql
                    .Execute
                End With
            cx.Close
        Set cx = Nothing
End Sub

Function AssignPullGroup(gNum As String) As Boolean

    Dim cn As ADODB.Connection
        Set cn = CreateObject("ADODB.Connection")
    Dim rs As ADODB.recordSet
        Set rs = CreateObject("ADODB.Recordset")
    
    cn.Open GetCString()
    rs.Open "tblGroups", cn, adLockOptimistic, adCmdTable
    
    strSQL = "UPDATE tblOAS SET assignedtoPuller = '" & formMicrofilm.tbPullAgent.Text & "' WHERE groupNum = '" & gNum & "'"
    cn.Execute strSQL
    strSQL = "UPDATE tblOAS SET pullAssignedB = True WHERE groupNum = '" & gNum & "'"
    cn.Execute strSQL
    
    rs.Close
    cn.Close
    
    cn.Open GetCString()
    rs.Open "tblGroups", cn, adLockOptimistic, adCmdTable
    
    With rs
        .Find "groupNum = '" & gNum & "'"
        .Fields("assignedToPuller") = formMicrofilm.tbPullAgent.Text
        .Update
    End With
    
    MsgBox gNum & " is assigned to you."

End Function

Function GetRecordSet(sql As String)

    Dim cn As ADODB.Connection
        Set cn = CreateObject("ADODB.Connection")
    
    Dim rs As ADODB.recordSet
        Set rs = CreateObject("ADODB.RECORDSET")
    
    cn.Open GetCString()
    rs.Open sql, cn, adLockOptimistic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        GetRecordSet = rs.GetRows
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Function

Function GetNewRecordSet(sql As String)
    
    Dim rs As ADODB.recordSet
        Set rs = New ADODB.recordSet
  
            rs.Open sql, GetCString(), adOpenKeyset
                Set GetNewRecordSet = rs.Clone
  
  rs.Close
  Set rs = Nothing
  
End Function

Function GetRecordSetOneParameter(sql As String, prm As ADODB.Parameter)
    
    Dim cn As ADODB.Connection
        Set cn = New ADODB.Connection
    
    Dim rs As ADODB.recordSet
        Set rs = CreateObject("ADODB.RECORDSET")
    
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
            cn.Open GetCString()
                With cmd
                    .ActiveConnection = cn
                    .CommandText = sql
                    .Parameters.Append prm
                End With
                
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        GetRecordSetOneParameter = rs.GetRows
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Function


Function GetNewRecordSetOneParameter(sql As String, prm As ADODB.Parameter)
    
    Dim cn As ADODB.Connection
        Set cn = New ADODB.Connection
        
    Dim rs As ADODB.recordSet
        Set rs = New ADODB.recordSet
  
    Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
            cn.Open GetCString()
                With cmd
                    .ActiveConnection = cn
                    .CommandText = sql
                    .Parameters.Append prm
                End With
                
            rs.Open cmd, , adOpenKeyset
                Set GetNewRecordSetOneParameter = rs.Clone
  
  rs.Close
  Set rs = Nothing
  
End Function
