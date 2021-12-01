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
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "三角形的面积"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "第三条边"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "第二条边"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "第一条边"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Dim a As Integer, b As Integer, c As Integer, t As Single
a = Int(Text1.Text)
b = Int(Text2.Text)
c = Int(Text3.Text)
t = (a + b + c) / 2
If a + b > c And a + c > b And b + c > a Then
    s = Sqr(t * (t - a) * (t - b) * (t - c))
    Text4.Text = s
Else
    MsgBox "您输入的三角形的边长都不符合要求"

End If


End Sub

