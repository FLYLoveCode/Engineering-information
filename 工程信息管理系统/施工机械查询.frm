VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 施工机械查询 
   Caption         =   "施工机械查询"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   14520
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   615
      Left            =   12720
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "刷新"
      Height          =   615
      Left            =   11160
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "同时满足两个条件查询"
      Height          =   1095
      Left            =   11400
      TabIndex        =   7
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "查询"
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "根据施工工程编号查询"
      Height          =   1695
      Left            =   11400
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "查询"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "根据施工机械名称查询"
      Height          =   1695
      Left            =   11400
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12938
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu fhzcd 
      Caption         =   "返回主菜单"
   End
End
Attribute VB_Name = "施工机械查询"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From 施工机械 Where 名称 = '" & Text1.Text & "'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command2_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From 施工机械 Where 工程编号 = '" & Text2.Text & "'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command3_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From 施工机械 Where 名称 = '" & Text1.Text & "' And 工程编号 = '" & Text2.Text & "'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command4_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From 施工机械"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub fhzcd_Click()
主菜单.Show
施工机械查询.Hide
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "施工机械"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

