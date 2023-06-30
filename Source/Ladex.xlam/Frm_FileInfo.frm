VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_FileInfo 
   Caption         =   "ファイル情報"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "Frm_FileInfo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim infoVal As String
  Dim objFileInfo As Object
  
  Const funcName As String = "Frm_FileInfo.UserForm_Initialize"

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
  
  'ファイル情報取得------------------------------
  Call Library.getFileInfo(ActiveWorkbook.path & "\" & ActiveWorkbook.Name, objFileInfo)
  infoVal = ""
  infoVal = infoVal & "シート数： " & Worksheets.count & vbNewLine
  infoVal = infoVal & "作成日時： " & objFileInfo("createAt") & vbNewLine
  infoVal = infoVal & "更新日時： " & objFileInfo("updateAt") & vbNewLine
  infoVal = infoVal & "サイズ　： " & Library.convscale(objFileInfo("size")) & " [" & Format(objFileInfo("size"), "#,##0") & " Byte" & "]" & vbNewLine
  
  
  With Me
    .Caption = "[" & thisAppName & "] ファイル情報 "
    .FilePath.Value = ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    .FileInfo.Value = infoVal
    
    
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
'キャンセル処理
Private Sub btnClose_Click()
  Unload Me
End Sub

