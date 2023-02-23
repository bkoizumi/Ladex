VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Wait 
   Caption         =   "手動確認"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "Frm_Wait.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'==================================================================================================
Private Sub UserForm_Initialize()
    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

    Dim ret As Long
    Dim formHWnd As Long

    'Get window handle of the userform
    formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)
    'If formHWnd = 0 Then Debug.Print Err.LastDllError

    'Set userform window to 'always on top'
    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'If ret = 0 Then Debug.Print Err.LastDllError

'    Application.WindowState = xlMinimized ' この操作は必須

End Sub




Private Sub btn_OK_Click()
  Call Ctl_Selenium.テスト結果(True, TextBox1.value)
  Unload Me
End Sub

Private Sub btn_NG_Click()
  Call Ctl_Selenium.テスト結果(False, TextBox1.value)
  Unload Me
End Sub

Private Sub Cancel_Click()
  Unload Me
  Call Library.errorHandle
  End
  
End Sub
