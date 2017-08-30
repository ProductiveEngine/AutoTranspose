Attribute VB_Name = "Utils"
Dim gconn As String

Public Enum CombineActionEnum
        Add = 1
        AddTime = 2
        Mean = 3
        MeanTime = 4
        MeanPercentage = 5
        WeightedMean = 6
        WeightedMeanTime = 7
        WeightedMeanPercentage = 8
End Enum

' 'Excel 12.0 Xml' for .xlsx files, 'Excel 12.0 Macro' for .xlsm and 'Excel 12.0' for .xlsb files
Public Function GetExcelConnection(ByVal path As String, _
    Optional ByVal headers As Boolean = True, Optional ByVal password As String = "", Optional ByVal xlsType As String = "") As String
    Dim strConn As String
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & path & ";" & _
               IIf(password = "", "", ";Jet OLEDB:Database Password=" & password & ";") & _
              "Extended Properties='Excel 12.0" & xlsType & ";HDR=" & _
              IIf(headers, "Yes", "No") & "'"
    
    GetExcelConnection = strConn
End Function

Public Function MakeConn(fullpath As String) As String
    gconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullpath & ";Persist Security Info=False"
    MakeConn = gconn
End Function

Public Function GetGConn() As String
    gconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\db.mdb;Persist Security Info=False"
    GetGConn = gconn
End Function

Public Function AdoConn()
    Dim con As New ADODB.Connection
    con.ConnectionString = GetGConn()
    con.CursorLocation = adUseClient
    
    AdoConn = con
End Function

Public Function DirectoryChooser(dialog As Control, initial As String) As String

    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    dialog.DialogTitle = "Select a directory" 'titlebar
    dialog.InitDir = App.path 'start dir, might be "C:\" or so also
    dialog.FileName = "Select a Directory"  'Something in filenamebox
    dialog.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    dialog.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    dialog.CancelError = True 'allow escape key/cancel
    dialog.ShowSave   'show the dialog screen

    If Err <> 32755 Then    ' User didn't chose Cancel.
        DirectoryChooser = CurDir
    Else
        DirectoryChooser = initial
    End If

    ChDir sTempDir  'restore path to what it was at entering
End Function

Public Sub RunSql(conn As ADODB.Connection, cmd As String)
    conn.Open
    conn.Execute (cmd)
    conn.Close
End Sub

Public Function SqlStr(str As String) As String
    SqlStr = "'" & str & "'"
End Function

Public Function HasPassword(password As String) As Boolean
    HasPassword = False
    
    If Not IsNull(password) Then
        If Len(Trim(password)) > 0 Then
            HasPassword = True
        End If
    End If
End Function

Public Function StrNullOrEmpty(str As String) As Boolean

    If IsNull(str) Then
        StrNullOrEmpty = True
    ElseIf Len(Trim(str)) = 0 Then
        StrNullOrEmpty = True
    Else
        StrNullOrEmpty = False
    End If
End Function

Public Function HMStoSec(s As String) As Long
        
    strHMS = Replace(s, " ", ":")
    HMStoSec = Split(strHMS, ":")(0) * 3600 + _
               Split(strHMS, ":")(1) * 60 + _
               Split(strHMS, ":")(2)
End Function

Public Function RemoveNewLine(s As String) As String
    s = Replace(s, ChrW(10), "")
    s = Replace(s, ChrW(13), "")
    RemoveNewLine = s
End Function

Public Function FieldExistsInRS( _
   ByRef rs As ADODB.Recordset, _
   ByVal fieldName As String)
   Dim fld As ADODB.field
    
   For Each fld In rs.Fields
      If fld.name = fieldName Then
         FieldExistsInRS = True
         Exit Function
      End If
   Next
    
   FieldExistsInRS = False
End Function

