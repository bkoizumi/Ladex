VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Progress 
   Caption         =   "処理中"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5820
   OleObjectBlob   =   "Frm_Progress.frx":0000
   StartUpPosition =   3  'Windows の既定値
End
Attribute VB_Name = "Frm_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const HWND_TOPMOST As Long = -1
'Private Const SWP_NOSIZE As Long = &H1&
'Private Const SWP_NOMOVE As Long = &H2&
'
'
'Private Sub UserForm_Activate()
'    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'    Me.StartUpPosition = 1
'End Sub



'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = vbFormControlMenu Then
'        Cancel = True
'    End If
'End Sub

'==================================================================================================
Private Sub UserForm_Initialize()
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "ProGressBar |  " & thisAppName

End Sub


