VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_GetFile 
   Caption         =   "�t�@�C���Ǘ�"
   ClientHeight    =   4690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235.001
   OleObjectBlob   =   "Frm_GetFile.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_GetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InitializeFlg   As Boolean
Public selectLine   As Long

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Dim i As Variant
  
  Const funcName As String = "Frm_GetFile.UserForm_Initialize"

  '�����J�n--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '�\���ʒu�w��----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True

  With Frm_GetFile
    .Caption = "[" & thisAppName & "] ���C�ɓ���"
  End With
  InitializeFlg = False
  
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
'�Q�ƃ{�^��[�f�B���N�g��]
Private Sub DirPath01_Click()
  Dim targetDir As String
  
  targetDir = Library.getDirPath(targetDir01.Value, , "getFileDirPath")
  If targetDir <> "" Then
    targetDir01.Value = targetDir
  End If
End Sub

'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
' ���s
Private Sub Submit_Click()
  
  FrmVal("targetDir01") = targetDir01.Value
  FrmVal("getFileName01") = getFileName01.Value
  FrmVal("getCreateAt01") = getCreateAt01.Value
  FrmVal("getUpdateAt01") = getUpdateAt01.Value
  FrmVal("getExtension01") = getExtension01.Value
  FrmVal("getSize01") = getSize01.Value
  FrmVal("getFileNameOnly01") = getFileNameOnly01.Value
  FrmVal("getFullPath01") = getFullPath01.Value
  FrmVal("getSubDir01") = getSubDir01.Value

  
  Unload Me
End Sub

