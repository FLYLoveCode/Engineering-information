VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "抽奖"
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "单击抽奖按钮开始抽奖"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize
Dim a As Integer
a = Int(180 * Rnd) + 1
Label1.Caption = "恭喜" & a & "号中奖了！"

End Sub

