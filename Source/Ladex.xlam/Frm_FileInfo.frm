VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_FileInfo 
   Caption         =   "�t�@�C�����"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "Frm_FileInfo.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_FileInfo"
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
  Dim infoVal As String
  Dim objFileInfo As Object
  
  Const funcName As String = "Frm_FileInfo.UserForm_Initialize"

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
  
  '�t�@�C�����擾------------------------------
  Call Library.getFileInfo(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, objFileInfo)
  infoVal = ""
  infoVal = infoVal & "�V�[�g���F " & Worksheets.count & vbNewLine
  infoVal = infoVal & "�쐬�����F " & objFileInfo("createAt") & vbNewLine
  infoVal = infoVal & "�X�V�����F " & objFileInfo("updateAt") & vbNewLine
  infoVal = infoVal & "�T�C�Y�@�F " & Library.convscale(objFileInfo("size")) & " [" & Format(objFileInfo("size"), "#,##0") & " Byte" & "]" & vbNewLine
  
  
  With Me
    .Caption = "[" & thisAppName & "] �t�@�C����� "
    .FilePath.Value = ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    .FileInfo.Value = infoVal
    
    
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
'�L�����Z������
Private Sub btnClose_Click()
  Unload Me
End Sub

