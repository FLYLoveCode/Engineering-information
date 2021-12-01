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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const e As Single = 2.71828
Private Sub Form_Click()
Dim x As Integer
x = InputBox("请输入x的值", "请输入")
If x >= 9 Then
    Print Sin(x) + x ^ 3 + 5
Else
    Print e ^ x + Int(x)
End If
End Sub

