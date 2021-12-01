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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "µÇÂ¼²âÊÔ½Ó¿Ú"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "µÇÂ½"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "ÃÜÂë"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ÕË»§"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "fly" And Text2.Text = "0000" Then
    a = MsgBox("µÇÂ¼³É¹¦", vbInformation + vbOKOnly, "ÃÜÂëµÇÂ¼")
Else
    a = MsgBox("ÃÜÂë´íÎó", vbOKCancle, "ÇëÖØÊÔ")
End If

End Sub

Private Sub Command2_Click()
Call accessÊý¾Ý¿âÁ¬½Ó
End Sub

