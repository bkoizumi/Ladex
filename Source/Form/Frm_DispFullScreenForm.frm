VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_DispFullScreenForm 
   Caption         =   "�S��ʉ���"
   ClientHeight    =   360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "Frm_DispFullScreenForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_DispFullScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
  
  Call Library.setRegistry("UserForm", "Zoom01Top", Me.Top)
  Call Library.setRegistry("UserForm", "Zoom01Left", Me.Left)
  
  Call Library.startScript
  Application.DisplayFullScreen = False
  Call Library.endScript
  
  Unload Me
End Sub
