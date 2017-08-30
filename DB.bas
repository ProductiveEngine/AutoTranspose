Attribute VB_Name = "DB"
Dim conn As ADODB.Connection

Public Sub InitializeConn()
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetGConn
End Sub

Public Sub DeleteActualService(id As Long)
    Dim cmd As String
    
    InitializeConn
    
    cmd = "DELETE FROM RawService WHERE ActualServiceId = " & id
    RunSql conn, cmd
    cmd = "DELETE FROM DataColumns WHERE ActualServiceId = " & id
    RunSql conn, cmd
    cmd = "DELETE FROM ActualService WHERE Id = " & id
    RunSql conn, cmd
    
End Sub


