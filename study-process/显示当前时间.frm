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
   Begin VB.CommandButton Command2 
      Caption         =   "当前时间"
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "当前日期"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "显示当前系统时间"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "显示当前系统日期"
      Height          =   495
      Left            =   0
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
Label3.Caption = Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日"
End Sub

Private Sub Command2_Click()
Label4.Caption = Hour(Now) & "时" & Minute(Now) & "分" & Second(Now) & "秒"
End Sub

