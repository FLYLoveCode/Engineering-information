VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4935
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "四舍五入"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "小数点后的位数"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "四舍五入的数据"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Double, b As Double, n As Integer
a = Val(Text1.Text)
n = Text2.Text
a = a * 10 ^ n
b = a - Int(a)
Text1.Text = IIf(b < 0.5, Int(a), Int(a) + 1) / (10 ^ n)

End Sub

