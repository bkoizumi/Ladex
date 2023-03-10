VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_mkQRCode 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Frm_mkQRCode.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_mkQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InitializeFlg  As Boolean
Public selectLine     As Long





'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  '�����J�n--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '�\���ʒu�w��----------------------------------
  StartUpPosition = 0
  top = ActiveWindow.top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  
  
  CellAddress.Text = Library.getColumnName(ActiveCell.Column + 1)
  codeSize.Text = 140
  
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
End Sub


'==================================================================================================
Private Sub OKButton_Click()
  
  FrmVal.add "CellAddress", CellAddress.Text
  FrmVal.add "codeSize", codeSize.Text
  FrmVal.add "onReSize", onReSize.Value
  
  Unload Me

End Sub


