VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Auto Transpose"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProgress"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRun"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtLog"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmbServices"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdRefresh"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Settings"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblRawDataFolder"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "lblServices"
      Tab(1).Control(4)=   "txtRawDataFolder"
      Tab(1).Control(5)=   "cmdRawData"
      Tab(1).Control(6)=   "CommonDialog1"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(8)=   "txtRawDataName"
      Tab(1).Control(9)=   "cmdAbout"
      Tab(1).Control(10)=   "cmdDeleteActualService"
      Tab(1).Control(11)=   "cmdAddActualService"
      Tab(1).Control(12)=   "lstActualServices"
      Tab(1).ControlCount=   13
      Begin VB.ListBox lstActualServices 
         Height          =   2400
         Left            =   -74685
         TabIndex        =   20
         Top             =   2925
         Width           =   3750
      End
      Begin VB.CommandButton cmdAddActualService 
         Caption         =   "Add"
         Height          =   465
         Left            =   -71640
         Picture         =   "Main.frx":0038
         TabIndex        =   19
         Top             =   5490
         Width           =   705
      End
      Begin VB.CommandButton cmdDeleteActualService 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -74685
         Picture         =   "Main.frx":538A
         TabIndex        =   18
         Top             =   5490
         Width           =   825
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   -68880
         TabIndex        =   17
         Top             =   5535
         Width           =   975
      End
      Begin VB.TextBox txtRawDataName 
         Height          =   375
         Left            =   -70230
         TabIndex        =   15
         Top             =   585
         Width           =   1875
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   330
         Left            =   2970
         TabIndex        =   13
         Top             =   585
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sheets"
         Height          =   1230
         Left            =   -74775
         TabIndex        =   7
         Top             =   1170
         Width           =   6900
         Begin VB.TextBox txtFinalReportSheet 
            Height          =   375
            Left            =   2055
            TabIndex        =   11
            Top             =   705
            Width           =   1815
         End
         Begin VB.TextBox txtRawDataSheet 
            Height          =   375
            Left            =   2055
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Final Report sheet"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   795
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Raw data sheet"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   315
            Width           =   1815
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -66840
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdRawData 
         Caption         =   "..."
         Height          =   315
         Left            =   -73485
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtRawDataFolder 
         Height          =   375
         Left            =   -72735
         TabIndex        =   5
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbServices 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   585
         Width           =   2415
      End
      Begin VB.TextBox txtLog 
         Height          =   1965
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2520
         Width           =   5010
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label lblServices 
         Caption         =   "Services for"
         Height          =   195
         Left            =   -74685
         TabIndex        =   21
         Top             =   2565
         Width           =   1140
      End
      Begin VB.Label Label6 
         Caption         =   ".xlsx"
         Height          =   255
         Left            =   -68250
         TabIndex        =   16
         ToolTipText     =   "Raw data filename"
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   -70725
         TabIndex        =   14
         ToolTipText     =   "Raw data filename"
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblProgress 
         Caption         =   "Progress"
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   2115
         Width           =   5235
      End
      Begin VB.Label lblRawDataFolder 
         Caption         =   "Raw data folder"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   645
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim statement As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim firstLoad As Boolean

Dim curDate As Date

Enum SettingsEnum
        RawDataFolder = 1
        RawDataSheet = 2
        NamedReportSheet = 3
        FinalReportSheet = 4
        RawDataName = 5
End Enum

Dim mSelectedActualServiceId As Long
Dim mSelectedActualServiceName As String
    
Dim mServiceToRun As Long
Dim mAllServices As New Collection

Private Sub cmbServices_Click()
    If cmbServices.ListIndex >= 0 Then
        mServiceToRun = cmbServices.ItemData(cmbServices.ListIndex)
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
    Unload frmAbout
End Sub

Private Sub cmdRefresh_Click()
    PopulateServices
End Sub

Private Sub cmdRun_Click()
    
    If mServiceToRun > 0 Then
        cmdRun.Enabled = False
        txtLog.Text = ""
     
        Dim i As Integer
     
        PopulateServices
        Transpose mServiceToRun, txtLog
     
        MsgBox "Process completed!"
        cmdRun.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    firstLoad = True
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetGConn
    
    PopulateServices
    PopulateActualServices
    PopulateSettings
    
    curDate = Now()
  firstLoad = False
End Sub

Private Sub OpenFrmActualService()
    frmActualService.Show vbModal, Me
    Unload frmActualService
    PopulateActualServices
End Sub
'----------------------------------------------------
'Populate
'----------------------------------------------------
Private Sub PopulateServices()
    Dim serviceId As Long
    
    conn.Open
    
    Set mAllServices = New Collection
    
    statement = "SELECT Id, ActualServiceName FROM ActualService ORDER BY ActualServiceName"
    Set rs = conn.Execute(statement, , adCmdText)
   
    cmbServices.Clear
    
    Do While Not rs.EOF
        cmbServices.AddItem rs!ActualServiceName
        serviceId = rs!id
        cmbServices.ItemData(cmbServices.NewIndex) = serviceId
        mAllServices.Add serviceId
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
        
End Sub

Private Sub PopulateSettings()
    conn.Open

    ' Select the data.
    statement = "SELECT Code, ActualValue FROM Settings"
    Set rs = conn.Execute(statement, , adCmdText)
    
    Do While Not rs.EOF
        If rs!Code = SettingsEnum.RawDataFolder Then
            txtRawDataFolder.Text = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.RawDataSheet Then
            txtRawDataSheet.Text = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.FinalReportSheet Then
            txtFinalReportSheet.Text = rs!ActualValue
        ElseIf rs!Code = SettingsEnum.RawDataName Then
            txtRawDataName.Text = rs!ActualValue
        End If
        rs.MoveNext
    Loop

    ' Close the recordset and connection.
    rs.Close
    conn.Close
End Sub

Private Sub PopulateActualServices()
    conn.Open
    
    statement = "SELECT Id, ActualServiceName FROM ActualService ORDER BY ActualServiceName"
    Set rs = conn.Execute(statement, , adCmdText)
    
    lstActualServices.Clear
    
    Do While Not rs.EOF
        lstActualServices.AddItem rs!ActualServiceName
        lstActualServices.ItemData(lstActualServices.NewIndex) = rs!id

        rs.MoveNext
    Loop

    rs.Close
    conn.Close
End Sub
'----------------------------------------------------
'Events
'----------------------------------------------------
Private Sub lstActualServices_Click()
    If lstActualServices.ListIndex >= 0 Then
        mSelectedActualServiceId = lstActualServices.ItemData(lstActualServices.ListIndex)
        mSelectedActualServiceName = lstActualServices.Text
    End If
End Sub

Private Sub lstActualServices_DblClick()
    If lstActualServices.ListIndex >= 0 Then
        frmActualService.mActualServiceId = lstActualServices.ItemData(lstActualServices.ListIndex)
    
        OpenFrmActualService
        lstActualServices.Clear
        PopulateActualServices
    End If
End Sub

Private Sub txtRawDataFolder_Change()
    Dim cmd As String
    
    cmd = "UPDATE Settings SET ActualValue = " & SqlStr(txtRawDataFolder.Text) & " WHERE Code = " & SettingsEnum.RawDataFolder
        
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub txtRawDataSheet_Change()
    Dim cmd As String
    
    cmd = "UPDATE Settings SET ActualValue = " & SqlStr(txtRawDataSheet.Text) & " WHERE Code = " & SettingsEnum.RawDataSheet
    
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub txtRawDataName_Change()
    Dim cmd As String
    
    cmd = "UPDATE Settings SET ActualValue = " & SqlStr(txtRawDataName.Text) & " WHERE Code = " & SettingsEnum.RawDataName
    
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub txtFinalReportSheet_Change()
    Dim cmd As String
    
    cmd = "UPDATE Settings SET ActualValue = " & SqlStr(txtFinalReportSheet.Text) & " WHERE Code = " & SettingsEnum.FinalReportSheet
    
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub cmdAddActualService_Click()
    Dim cmd As String
    
    cmd = "INSERT INTO ActualService(ActualServiceName) Values ('NEW')"
    RunSql conn, cmd
        
    PopulateActualServices
End Sub

Private Sub cmdDeleteActualService_Click()
    If lstActualServices.ListIndex >= 0 Then
        If MsgBox("Are you sure you want to delete  " & mSelectedActualServiceName & " ?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
            DeleteActualService mSelectedActualServiceId
            PopulateActualServices
        End If
    End If
End Sub

Private Sub cmdRawData_Click()
    txtRawDataFolder.Text = DirectoryChooser(CommonDialog1, txtRawDataFolder.Text)
End Sub


