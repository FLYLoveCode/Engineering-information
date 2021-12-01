VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "长度"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "小写"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "大写"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "请输入英文字母"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Command2_Click()
Text1.Text = LCase(Text1.Text)
End Sub

Private Sub Command3_Click()
Text2.Text = Len(Text1.Text)
End Sub
