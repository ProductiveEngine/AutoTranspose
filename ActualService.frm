VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmActualService 
   Caption         =   "Service Settings"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   11324
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Data"
      TabPicture(0)   =   "ActualService.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtActualServiceName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dgDataColumns"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtLegend"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "AdodcDataColumns"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CommonDialog1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtOutputName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOutput"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtOutputFolder"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Services"
      TabPicture(1)   =   "ActualService.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "AdodcServices"
      Tab(1).Control(1)=   "dgServices"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtOutputFolder 
         Height          =   375
         Left            =   2430
         TabIndex        =   9
         Top             =   945
         Width           =   1590
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "..."
         Height          =   330
         Left            =   1845
         TabIndex        =   8
         Top             =   990
         Width           =   420
      End
      Begin VB.TextBox txtOutputName 
         Height          =   375
         Left            =   4770
         TabIndex        =   7
         Top             =   945
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7695
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc AdodcDataColumns 
         Height          =   330
         Left            =   7650
         Top             =   5625
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Data Column"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtLegend 
         Enabled         =   0   'False
         Height          =   2490
         Left            =   7650
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1980
         Width           =   3120
      End
      Begin MSDataGridLib.DataGrid dgDataColumns 
         Height          =   4335
         Left            =   270
         TabIndex        =   4
         Top             =   1620
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "InitialName"
            Caption         =   "Initial Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "OutputName"
            Caption         =   "Output Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "CombineAction"
            Caption         =   "Combine Action"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "WeightColumn"
            Caption         =   "Weight Column"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "ActualServiceId"
            Caption         =   "ActualServiceId"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtActualServiceName 
         Height          =   375
         Left            =   1845
         TabIndex        =   3
         Top             =   495
         Width           =   5550
      End
      Begin MSDataGridLib.DataGrid dgServices 
         Bindings        =   "ActualService.frx":0038
         Height          =   5265
         Left            =   -74775
         TabIndex        =   1
         Top             =   675
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9287
         _Version        =   393216
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "ServiceName"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ActualServiceId"
            Caption         =   "ActualServiceId"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3270.047
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdodcServices 
         Height          =   375
         Left            =   -67260
         Top             =   585
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Rows"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Caption         =   "Output Folder"
         Height          =   285
         Left            =   225
         TabIndex        =   12
         Top             =   1035
         Width           =   1545
      End
      Begin VB.Label Label10 
         Caption         =   "Name"
         Height          =   285
         Left            =   4185
         TabIndex        =   11
         ToolTipText     =   "Output report filename"
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   ".xlsx"
         Height          =   240
         Left            =   6300
         TabIndex        =   10
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Combine Action Legend"
         Height          =   240
         Left            =   7650
         TabIndex        =   5
         Top             =   1665
         Width           =   3075
      End
      Begin VB.Label Label3 
         Caption         =   "Actual Service Name"
         Height          =   285
         Left            =   225
         TabIndex        =   2
         Top             =   585
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmActualService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim statement As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Public mActualServiceId As Long

Dim firstLoad As Boolean

Private Sub Form_Load()
    firstLoad = True
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetGConn
    conn.CursorLocation = adUseClient
    
    PopulateBasicSettings
    FillServices
    FillDataColumns
    
    txtLegend.Text = "1: Add" & vbCrLf & _
        "2: Add time" & vbCrLf & _
        "3: Mean" & vbCrLf & _
        "4: Mean time" & vbCrLf & _
        "5: Mean percentage" & vbCrLf & _
        "6: Weighted Mean" & vbCrLf & _
        "7: Weighted Mean time" & vbCrLf & _
        "8: Weighted Mean percentage" & vbCrLf & _
        "Any other value will de ignored"

    firstLoad = False
End Sub

Private Sub PopulateBasicSettings()
    conn.Open
    statement = "SELECT * FROM ActualService WHERE Id = " & mActualServiceId
    Set rs = conn.Execute(statement, , adCmdText)
    
    Do While Not rs.EOF
        txtActualServiceName.Text = IIf(IsNull(rs!ActualServiceName), "", rs!ActualServiceName)
        txtOutputFolder.Text = IIf(IsNull(rs!OutputFolder), "", rs!OutputFolder)
        txtOutputName.Text = IIf(IsNull(rs!OutputName), "", rs!OutputName)
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
End Sub

Private Sub FillServices()
    
    conn.Open
    AdodcServices.ConnectionString = conn.ConnectionString
    AdodcServices.RecordSource = "SELECT * FROM RawService WHERE ActualServiceId = " & mActualServiceId
    
    Set dgServices.DataSource = AdodcServices
    dgServices.Refresh
    conn.Close
End Sub

Private Sub FillDataColumns()
    
    conn.Open
    AdodcDataColumns.ConnectionString = conn.ConnectionString
    AdodcDataColumns.RecordSource = "SELECT * FROM DataColumns WHERE ActualServiceId = " & mActualServiceId
    
    Set dgDataColumns.DataSource = AdodcDataColumns
    dgDataColumns.Refresh
    conn.Close
End Sub

'----------------------------------------------------
'Events
'----------------------------------------------------
Private Sub txtOutputName_Change()
    Dim cmd As String
    
    cmd = "UPDATE ActualService SET OutputName = " & SqlStr(txtOutputName.Text) & " WHERE Id = " & mActualServiceId
        
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub txtActualServiceName_Change()
    Dim cmd As String
    
    cmd = "UPDATE ActualService SET ActualServiceName = " & SqlStr(txtActualServiceName.Text) & " WHERE Id = " & mActualServiceId
        
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub txtOutputFolder_Change()
    Dim cmd As String
    
    cmd = "UPDATE ActualService SET OutputFolder = " & SqlStr(txtOutputFolder.Text) & " WHERE Id = " & mActualServiceId
        
    If Not firstLoad Then
        RunSql conn, cmd
    End If
End Sub

Private Sub cmdOutput_Click()
    txtOutputFolder.Text = DirectoryChooser(CommonDialog1, txtOutputFolder.Text)
End Sub

Private Sub dgDataColumns_OnAddNew()
    dgDataColumns.Columns(4).Value = mActualServiceId
End Sub

Private Sub dgServices_OnAddNew()
    dgServices.Columns(1).Value = mActualServiceId
End Sub
