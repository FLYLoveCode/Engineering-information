VERSION 5.00
Begin VB.Form ���˵� 
   Caption         =   "���˵�"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   Picture         =   "���˵�.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   9495
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu zgxx 
      Caption         =   "ְ����Ϣ"
      Begin VB.Menu zgxxcx 
         Caption         =   "ְ����Ϣ��ѯ"
      End
      Begin VB.Menu zgxxgl 
         Caption         =   "ְ����Ϣ����"
      End
   End
   Begin VB.Menu dqsb 
      Caption         =   "�����豸"
      Begin VB.Menu dqsbcx 
         Caption         =   "�����豸��ѯ"
      End
      Begin VB.Menu dqsbgl 
         Caption         =   "�����豸����"
      End
   End
   Begin VB.Menu sgjx 
      Caption         =   "ʩ����е"
      Begin VB.Menu sgjxcx 
         Caption         =   "ʩ����е��ѯ"
      End
      Begin VB.Menu sgjxgl 
         Caption         =   "ʩ����е����"
      End
   End
   Begin VB.Menu qxgl 
      Caption         =   "Ȩ�޹���"
   End
   Begin VB.Menu tc 
      Caption         =   "�˳�"
   End
End
Attribute VB_Name = "���˵�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dqsbcx_Click()
�����豸��ѯ.Show
End Sub

Private Sub dqsbgl_Click()
�����豸����.Show
End Sub

Private Sub Form_Load()
��¼����.Hide

If ��¼����.Text1.Text = "199094061" Then
qxgl.Enabled = True
dqsbcx.Enabled = True
sgjxcx.Enabled = True
zgxxcx.Enabled = True
dqsbgl.Enabled = True
sgjxgl.Enabled = True
zgxxgl.Enabled = True

ElseIf ��¼����.Text1.Text = "199094041" Or ��¼����.Text1.Text = "199094057" Then
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
Ȩ�޹���.Show
End Sub

Private Sub sgjxcx_Click()
ʩ����е��ѯ.Show
End Sub

Private Sub sgjxgl_Click()
ʩ����е����.Show
End Sub

Private Sub tc_Click()
End

End Sub

Private Sub zgxxcx_Click()
ְ����Ϣ��ѯ.Show
End Sub

Private Sub zgxxgl_Click()
ְ����Ϣ����.Show
End Sub
