VERSION 5.00
Begin VB.Form 主菜单 
   Caption         =   "主菜单"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   Picture         =   "主菜单.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   9495
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu zgxx 
      Caption         =   "职工信息"
      Begin VB.Menu zgxxcx 
         Caption         =   "职工信息查询"
      End
      Begin VB.Menu zgxxgl 
         Caption         =   "职工信息管理"
      End
   End
   Begin VB.Menu dqsb 
      Caption         =   "电气设备"
      Begin VB.Menu dqsbcx 
         Caption         =   "电气设备查询"
      End
      Begin VB.Menu dqsbgl 
         Caption         =   "电气设备管理"
      End
   End
   Begin VB.Menu sgjx 
      Caption         =   "施工机械"
      Begin VB.Menu sgjxcx 
         Caption         =   "施工机械查询"
      End
      Begin VB.Menu sgjxgl 
         Caption         =   "施工机械管理"
      End
   End
   Begin VB.Menu qxgl 
      Caption         =   "权限管理"
   End
   Begin VB.Menu tc 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "主菜单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dqsbcx_Click()
电气设备查询.Show
End Sub

Private Sub dqsbgl_Click()
电气设备管理.Show
End Sub

Private Sub Form_Load()
登录界面.Hide

If 登录界面.Text1.Text = "199094061" Then
qxgl.Enabled = True
dqsbcx.Enabled = True
sgjxcx.Enabled = True
zgxxcx.Enabled = True
dqsbgl.Enabled = True
sgjxgl.Enabled = True
zgxxgl.Enabled = True

ElseIf 登录界面.Text1.Text = "199094041" Or 登录界面.Text1.Text = "199094057" Then
dqsbcx.Enabled = True
sgjxcx.Enabled = True
zgxxcx.Enabled = True
dqsbgl.Enabled = True
sgjxgl.Enabled = True
zgxxgl.Enabled = True
qxgl.Enabled = False



Else
sgjxcx.Enabled = True
dqsbcx.Enabled = True
zgxxcx.Enabled = True
dqsbgl.Enabled = False
sgjxgl.Enabled = False
zgxxgl.Enabled = False
qxgl.Enabled = False
End If
End Sub

Private Sub qxgl_Click()
权限管理.Show
End Sub

Private Sub sgjxcx_Click()
施工机械查询.Show
End Sub

Private Sub sgjxgl_Click()
施工机械管理.Show
End Sub

Private Sub tc_Click()
End

End Sub

Private Sub zgxxcx_Click()
职工信息查询.Show
End Sub

Private Sub zgxxgl_Click()
职工信息管理.Show
End Sub
