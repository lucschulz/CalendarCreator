VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   Designed by Luc Schulz
'   March 22, 2018
'
'   GetConnection() returns a database connection object using the path and password supplied in the properties below.
'   The AppVersion() property can be used to validate the current workbook's version against the version set in the database and in the UI.
'   The path and password to the database are kept in private properties inside a class module to limit their exposure in memory.


''' THE CURRENT RELEASE VERSION OF THE APPLICATION
Public Property Get AppVersion() As String
    AppVersion = "1.0.0"
End Property

''' RETURNS A CONNECTION TO THE DATABASE
Public Function GetConnection() As ADODB.Connection
    Dim cn As New ADODB.Connection
        cn.ConnectionString = ConnectionString()
            Set GetConnection = cn
End Function

''' RETURNS A CONNECTION STRING FOR CONNECTION TO THE DATABASE
Private Property Get ConnectionString() As String
    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabasePath() & ";Jet OLEDB:Database Password=" & DatabasePassword() & ";"
End Property

''' RETURNS THE CURRENT USER'S WINDOWS USERNAME
Private Property Get AgentUsername() As String
    AgentUsername = VBA.Environ("USERNAME")
End Property

''' RETURNS THE FULLY QUALIFIED PATH TO THE DATABASE FILE
Private Property Get DatabasePath() As String
    DatabasePath = ""
End Property

''' RETURNS THE DATABASE PASSWORD (ONLY STORE IN THIS LOCATION)
Private Property Get DatabasePassword() As String
    DatabasePassword = ""
End Property




''''''''''''''''''''''''''''''''''''''''''''
''' ACCESS LOG
''''''''''''''''''''''''''''''''''''''''''''
Public Sub WriteAccessLogEntry(action As String)
    
    On Error GoTo Exception
    
    Dim prms As New cSqlParameters
    
    Dim cn As ADODB.Connection
    Set cn = GetConnection()
    cn.Open
    
    Dim cmd As New ADODB.Command
    Dim prmUser As ADODB.Parameter
    Set prmUser = prms.NewVarCharParameter("@Username", 20, Environ("USERNAME"))
    
    Dim prmVersion As ADODB.Parameter
    Set prmVersion = prms.NewVarCharParameter("@Version", 10, AppVersion())
    
    Dim prmUserAction As ADODB.Parameter
    Set prmUserAction = prms.NewVarCharParameter("@UserAction", 20, action)
    
    With cmd
        .ActiveConnection = cn
        .CommandText = "INSERT INTO AccessLog(Username, LastAccess, Version, UserAction) VALUES(@Username, NOW, @Version, @UserAction);"
        .Parameters.Append prmUser
        .Parameters.Append prmVersion
        .Parameters.Append prmUserAction
        .Execute
    End With
    
    cn.Close
    
    Set cmd = Nothing
    Set cn = Nothing
    
    Exit Sub
    
Exception:
    MsgBox "Unable to log access to database. Press OK to continue.", vbOKOnly + vbInformation, "Unable to Log"
    
End Sub















