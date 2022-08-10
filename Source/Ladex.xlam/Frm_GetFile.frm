VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_GetFile 
   Caption         =   "ファイル管理"
   ClientHeight    =   4690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235.001
   OleObjectBlob   =   "Frm_GetFile.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Dim i As Variant
  
  Const funcName As String = "Frm_GetFile.UserForm_Initialize"

  '処理開始--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True

  With Frm_GetFile
    .Caption = "[" & thisAppName & "] お気に入り"
  End With
  InitializeFlg = False
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub



'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'参照ボタン[ディレクトリ]
Private Sub DirPath01_Click()
  Dim targetDir As String
  
  targetDir = Library.getDirPath(targetDir01.Value, , "getFileDirPath")
  If targetDir <> "" Then
    targetDir01.Value = targetDir
  End If
End Sub

'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
' 実行
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

