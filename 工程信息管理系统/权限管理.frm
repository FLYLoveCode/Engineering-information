VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 权限管理 
   Caption         =   "权限管理"
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   9435
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "保存"
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入用户名查询"
      Height          =   2295
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4260
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
      Height          =   495
      Left            =   1200
      Top             =   1080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
Attribute VB_Name = "权限管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from 管理员 where 用户名 ='"
While Not Adodc1.Recordset.EOF
Combo1.AddItem Adodc1.Recordset.Fields
End With

End Sub

Private Sub Command1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From 管理员 Where 用户名 = '" & Text1.Text & "'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command2_Click()
a = MsgBox("添加成功", vbInformation + vbOKOnly, "提示")

Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
a = MsgBox("是否删除该用户", vbokcancle, "警告")


Adodc1.Recordset.Delete
a = MsgBox("删除成功", vbInformation + vbOKOnly, "提示")

Adodc1.Refresh
End Sub

Private Sub Command4_Click()
a = MsgBox("保存成功", vbInformation + vbOKOnly, "提示")
Adodc1.Recordset.Update
End Sub

Private Sub fhzcd_Click()
主菜单.Show
权限管理.Hide
End Sub

Private Sub Form_Load()

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\数据库.mdb;Persist Security Info=False"
Adodc1.RecordSource = "管理员"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Label3_Click()

End Sub

