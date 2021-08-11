VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_DispFullScreenForm 
   Caption         =   "全画面解除"
   ClientHeight    =   360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "Frm_DispFullScreenForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_DispFullScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
  
  Call Library.setRegistry("UserForm", "Zoom01Top", Me.Top)
  Call Library.setRegistry("UserForm", "Zoom01Left", Me.Left)
  
  Application.DisplayFullScreen = False
  Unload Me
End Sub
