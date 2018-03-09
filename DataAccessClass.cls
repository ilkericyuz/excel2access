VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccessClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Public conn As ADODB.connection
    Public cmd As ADODB.Command
    Public rs As ADODB.Recordset
    
    
Private Sub class_initialize()
    ''' Sets connection, command and recordset
    
    Set conn = New ADODB.connection
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset
    
End Sub


Private Sub Class_Terminate()
    ''' Resets connection, command and recordset
    
    Set conn = Nothing
    Set cmd = Nothing
    Set rs = Nothing

End Sub


Public Sub ConnectToDB(theDbName As String, theDbPath As String)
    ''' Sets connection to MS Access database according to the parameters.
    '''
    ''' @input:
    '''     theDbName:      Name of the MS Access database file with extension.
    '''     theDbPath:      Path of the MS Access database folder with backslash at the end.
    '''
    ''' @output: None
    
    ' Set connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" _
                                    & "Data Source=" & theDbPath & theDbName & ";" _
                                    & "Persist Security Info=False;"
    conn.Open
    cmd.ActiveConnection = conn
    
End Sub


Public Sub DisconnectFromDB()
    ''' Disconnects from the connected database and resets variables
    '''
    ''' @input: None
    '''
    ''' @output: None
    
    On Error Resume Next
    
    conn.Close
    Call Class_Terminate

End Sub

