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
   Begin VB.CommandButton Command1 
      Caption         =   "显示出生日期"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "请输入身份证号码"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1.Text) <> 18 Then
    MsgBox "您输入的身份证号码有误"
Else
    Text2.Text = Mid(Text1.Text, 7, 4) & "年 " & Mid(Text1.Text, 11, 2) & "月" & Mid(Text1.Text, 13, 2) & "日"
End If
    

End Sub

Private Sub Label2_Click()

End Sub

