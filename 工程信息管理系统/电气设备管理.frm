VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form �����豸���� 
   Caption         =   "�����豸����"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   13365
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   10200
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      Height          =   615
      Left            =   7560
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ɾ��"
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����������Ϣ��ѯ"
      Height          =   3975
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "ˢ��"
         Height          =   615
         Left            =   2640
         TabIndex        =   7
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ѯ"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "���̱��"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   1440
      Top             =   2520
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
         Name            =   "����"
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
      Caption         =   "�������˵�"
   End
End
Attribute VB_Name = "�����豸����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\���ݿ�.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From �����豸 Where ���� = '" & Text1.Text & "' And ���̱�� = '" & Text2.Text & "'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command2_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\���ݿ�.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * From ʩ����е"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command3_Click()
a = MsgBox("��ӳɹ�", vbInformation + vbOKOnly, "��ʾ")

Adodc1.Recordset.AddNew
End Sub

Private Sub Command4_Click()
a = MsgBox("�Ƿ�ɾ�����豸", vbokcancle, "����")


Adodc1.Recordset.Delete
a = MsgBox("ɾ���ɹ�", vbInformation + vbOKOnly, "��ʾ")

Adodc1.Refresh
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command6_Click()
End

End Sub

Private Sub fhzcd_Click()
���˵�.Show
�����豸����.Hide
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\���ݿ�.mdb;Persist Security Info=False"
Adodc1.RecordSource = "�����豸"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
End Sub
