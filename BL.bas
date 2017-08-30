Attribute VB_Name = "BL"
Option Explicit

Dim excelApp As Excel.Application
Dim excelWB As Excel.Workbook
Dim excelWS As Excel.Worksheet
    
Dim statement As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

'Settings
Dim mRawDataFolder As String
Dim mRawDataSheet As String
Dim mFinalReportSheet As String
Dim mRawDataName As String

'Actual Service Settings
Dim mActualServiceName As String
Dim mActualServiceId As Long
Dim mOutputFolder As String
Dim mOutputName As String

Dim mLog As Control

Public Sub InitializeConn()
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetGConn
End Sub

Public Sub Transpose(currentServiceId As Long, txtLog As Control)
    On Error GoTo ErrorHandler
    Set mLog = txtLog
    
    GetSettings
    GetActualService currentServiceId
    
    Set excelApp = CreateObject("Excel.Application")
    'TODO: Check files integrity
        
    'TODO: if 0...
    Dim serviceNames As String
    serviceNames = GetServiceNames(currentServiceId)
    
    'Open Raw data file
    Dim rawConn As New ADODB.Connection
    Dim rsRaw As New ADODB.Recordset
    Dim rawString As String
    
    rawString = GetExcelConnection(mRawDataFolder & "\" & mRawDataName & ".xlsx")
    statement = "SELECT * FROM [" & mRawDataSheet & "$] WHERE ServiceId IN (" & serviceNames & ")"
        
    rawConn.Open rawString
    rsRaw.Open statement, rawConn, 1, 3
    
    If rsRaw.RecordCount > 0 Then
        Dim dataColumns As New Collection
        Dim data As New Collection
        
        Set dataColumns = GetDataColumns(currentServiceId)
        'datacolumns > 0...
        
        'Process Raw Data
        If rsRaw.RecordCount > 1 Then
            Set data = ScanMultipleRawRows(rsRaw, dataColumns)
        Else
            Set data = ScanSingleRawRow(rsRaw, dataColumns)
        End If
        
        rsRaw.Close
        rawConn.Close
        '-----------------------------------------------------
        UpdateData data
    Else
        'TODO: if 0...
    End If
    
    Excel.Application.Quit
    Set excelApp = Nothing
        
    mActualServiceId = 0
    mLog.Text = mLog.Text & "OK: " & mActualServiceName & vbCrLf
    mLog.Text = mLog.Text & "---------------------" & vbCrLf
    
    Exit Sub
ErrorHandler:
    mLog.Text = mLog.Text & "ERROR: " & mActualServiceName & vbCrLf & Err.Description & vbCrLf
    mLog.Text = mLog.Text & "---------------------" & vbCrLf
    
End Sub
'-------------------------------
'Load Data
'-------------------------------
Private Sub GetSettings()
    InitializeConn
    conn.Open
        
    statement = "SELECT Code, ActualValue FROM Settings"
    Set rs = conn.Execute(statement, , adCmdText)
    
    Do While Not rs.EOF
        If rs!Code = SettingsEnum.RawDataFolder Then
            mRawDataFolder = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.RawDataSheet Then
            mRawDataSheet = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.FinalReportSheet Then
            mFinalReportSheet = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.RawDataName Then
            mRawDataName = rs!ActualValue
        End If
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
End Sub

Private Sub GetActualService(id As Long)
    InitializeConn
    conn.Open
    
    statement = "SELECT * FROM ActualService WHERE Id = " & id
    Set rs = conn.Execute(statement, , adCmdText)
       
    Do While Not rs.EOF
        mActualServiceId = rs!id
        mActualServiceName = rs!ActualServiceName
        mOutputFolder = rs!OutputFolder
        mOutputName = rs!OutputName
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
End Sub

Private Function GetDataColumns(actualServiceId As Long) As Collection
    Dim i
    Dim column As DataColumn
    Dim dataColumns As New Collection
        
    InitializeConn
    conn.Open
        
    statement = "SELECT * FROM DataColumns WHERE ActualServiceId = " & actualServiceId
    Set rs = conn.Execute(statement, , adCmdText)
    
    Do While Not rs.EOF
        Set column = New DataColumn
        column.InitialName = rs!InitialName
        column.OutputName = IIf(IsNull(rs!OutputName), "", rs!OutputName)
        column.CombineAction = IIf(IsNull(rs!CombineAction), 0, rs!CombineAction)
        column.WeightColumn = IIf(IsNull(rs!WeightColumn), "", rs!WeightColumn)
        
        dataColumns.Add column, column.InitialName
        
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    Set GetDataColumns = dataColumns
End Function
Private Function GetServiceNames(serviceId As Long) As String
    Dim serviceNames As String
    Dim hasValues As Boolean
    InitializeConn
    conn.Open
    
    hasValues = False
        
    statement = "SELECT ServiceName FROM RawService WHERE ActualServiceId = " & serviceId
    Set rs = conn.Execute(statement, , adCmdText)
            
    Do While Not rs.EOF
        serviceNames = rs!serviceName & "," & serviceNames
        hasValues = True
        rs.MoveNext
    Loop
    
    If Not hasValues Then
        rs.Close
        conn.Close
        GetServiceNames = ""
        Exit Function
    End If
    
    serviceNames = RemoveNewLine(Mid(serviceNames, 1, Len(serviceNames) - 1))
    
    rs.Close
    conn.Close
    GetServiceNames = serviceNames
End Function
'-------------------------------
'Excel Utils
'-------------------------------
Private Function OpenWorkbook(path As String, name As String, password As String) As Excel.Workbook
    If HasPassword(password) Then
        Set OpenWorkbook = excelApp.Workbooks.Open(path & "\" & name, password:=password)
    Else
        Set OpenWorkbook = excelApp.Workbooks.Open(path & "\" & name)
    End If
End Function

Private Function CheckWorkbookIndex(worksheets As Variant, name As String) As Integer
    Dim i As Integer
    
    CheckWorkbookIndex = -1
    
    For i = 1 To worksheets.Count
        If worksheets(i).name = name Then
            CheckWorkbookIndex = i
        End If
    Next i

End Function

'-------------------------------
'Main Process
'-------------------------------
Private Function ScanSingleRawRow(rs As ADODB.Recordset, dataColumns As Variant) As Variant
    Dim resultData As New resultData
    Dim data As New Collection
    Dim dc As DataColumn
    
    For Each dc In dataColumns
        Dim Key As String
        Key = IIf(StrNullOrEmpty(dc.OutputName), dc.InitialName, dc.OutputName)
        
        If FieldExistsInRS(rs, dc.InitialName) Then
            Set resultData = New resultData
            resultData.Value = rs.Fields(dc.InitialName).Value
            resultData.Key = Key
            data.Add resultData
        Else
        '.......
        End If
    Next dc
    
    Set ScanSingleRawRow = data
End Function

Private Function ScanSingleRawRowM(rs As ADODB.Recordset, dataColumns As Variant) As Variant
    Dim resultData As New resultData
    Dim data As New Collection
    Dim dc As DataColumn
    Dim field As Double
    Dim Weight As Double
    Dim columnName As String
    
    'Check if column exists
    For Each dc In dataColumns
        Dim Key As String
        field = 0
        Weight = 0
        
        columnName = RemoveNewLine(RemoveNewLine(dc.InitialName))
        
        If Not FieldExistsInRS(rs, columnName) Then
            '...
        End If
        
        'Main Field
        If CombineActionEnum.Add = dc.CombineAction Or CombineActionEnum.Mean = dc.CombineAction Or CombineActionEnum.MeanPercentage = dc.CombineAction Then
            field = rs.Fields(columnName).Value
        ElseIf CombineActionEnum.WeightedMean = dc.CombineAction Or CombineActionEnum.WeightedMeanPercentage = dc.CombineAction Then
            field = rs.Fields(columnName).Value * rs.Fields(dc.WeightColumn).Value
        ElseIf CombineActionEnum.AddTime = dc.CombineAction Or CombineActionEnum.MeanTime = dc.CombineAction Then
            field = HMStoSec(rs.Fields(columnName))
        ElseIf CombineActionEnum.WeightedMeanTime = dc.CombineAction Then
            field = HMStoSec(rs.Fields(columnName).Value) * rs.Fields(dc.WeightColumn).Value
        Else
            field = 0
        End If
                        
        'Weight Field
        If CombineActionEnum.WeightedMean = dc.CombineAction Or _
            CombineActionEnum.WeightedMeanPercentage = dc.CombineAction Or _
            CombineActionEnum.WeightedMeanTime = dc.CombineAction Then
               Weight = rs.Fields(dc.WeightColumn).Value
        End If
        
        Key = IIf(StrNullOrEmpty(dc.OutputName), columnName, dc.OutputName)
        
        Set resultData = New resultData
        resultData.Value = field
        resultData.Key = Key
        
        resultData.Weight = Weight
        Set resultData.ColumnData = dc
            
        data.Add resultData
    Next dc
    
    Set ScanSingleRawRowM = data
End Function

Private Function ScanMultipleRawRows(rs As ADODB.Recordset, dataColumns As Variant) As Variant
    Dim data As New Collection
    Dim resultData As New resultData
    Dim finalData As New Collection
    Dim totalColumns As Integer
    Dim mulRows As New Collection
    Dim mR As New Collection
    Dim i As Integer
    Dim Key As String
    
    totalColumns = dataColumns.Count
    
    rs.MoveFirst
            
    Do While Not rs.EOF
        Set data = New Collection
        Set data = ScanSingleRawRowM(rs, dataColumns)
        mulRows.Add data
        rs.MoveNext
    Loop
    
    Dim accumulateValue As Double
    Dim accumulateWeight As Double
    Dim columnTotal As Double
    
    Dim headAction As Integer
    Dim headInitialName As String
    Dim headOutputName As String
    Dim runOnce As Boolean
    Dim rowCount As Integer
    rowCount = mulRows.Count
        
    'Check!!!
    
    For i = 1 To mulRows(1).Count
        accumulateValue = 0
        accumulateWeight = 0
        
        For Each mR In mulRows
            'Head
            headAction = mR(i).ColumnData.CombineAction
            headInitialName = mR(i).ColumnData.InitialName
            headOutputName = mR(i).ColumnData.OutputName
            '----------------------
            If CombineActionEnum.WeightedMean = headAction Or _
                CombineActionEnum.WeightedMeanPercentage = headAction Or _
                CombineActionEnum.WeightedMeanTime = headAction Then
                
                    accumulateWeight = accumulateWeight + mR(i).Weight
            End If
            
            accumulateValue = accumulateValue + mR(i).Value
        Next mR
                    
        If CombineActionEnum.WeightedMean = headAction Or _
            CombineActionEnum.WeightedMeanPercentage = headAction Or _
            CombineActionEnum.WeightedMeanTime = headAction Then
            
            If accumulateWeight > 0 Then
                columnTotal = Round(accumulateValue / accumulateWeight, 2)
            Else
                columnTotal = 0
            End If
        ElseIf CombineActionEnum.Mean = headAction Or _
            CombineActionEnum.MeanPercentage = headAction Or _
            CombineActionEnum.MeanTime = headAction Then
            
            columnTotal = Round(accumulateValue / rowCount, 2)
        Else
            columnTotal = Round(accumulateValue, 2)
        End If
                                    
        Key = IIf(StrNullOrEmpty(headOutputName), headInitialName, headOutputName)
        
        Set resultData = New resultData
        
        resultData.Key = Key
        
        If CombineActionEnum.AddTime = headAction Or _
            CombineActionEnum.MeanTime = headAction Or _
            CombineActionEnum.WeightedMeanTime = headAction Then
                resultData.Value = Format(DateAdd("s", columnTotal, "00:00:00"), "hh:mm:ss")
        Else
            resultData.Value = columnTotal
        End If
        
        finalData.Add resultData
    Next i

    Set ScanMultipleRawRows = finalData
End Function

Private Sub UpdateData(data As Collection)

    Dim outputConn As New ADODB.Connection
    Dim rsOutput As New ADODB.Recordset
    Dim outputString As String
    Dim rw As Range
    Dim rd As New resultData
    '----------------------------
    Set excelWB = OpenWorkbook(mOutputFolder, mOutputName, "")
    
    Dim workSheetIndex As Integer
    workSheetIndex = CheckWorkbookIndex(excelWB.worksheets, mFinalReportSheet)
                
    Set excelWS = excelWB.worksheets(workSheetIndex)
            
    For Each rd In data
        For Each rw In excelWS.Rows
            If rw.Cells(1, 1).Value = "" Then
                Exit For
            End If
                        
            If rw.Cells(1, 1) = RemoveNewLine(rd.Key) Then
                rw.Cells(1, 2) = rd.Value
                Exit For
            End If
        Next rw
    Next rd
       
    excelWB.Save
    excelWB.Close
End Sub
