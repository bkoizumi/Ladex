Attribute VB_Name = "Ctl_Book"
Option Explicit

'**************************************************************************************************
' * ブック管理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 名前定義削除()
  Dim wb As Workbook, tmp As String
  funcName = "Ctl_Book.名前定義削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each wb In Workbooks
    Workbooks(wb.Name).Activate
    Call Library.delVisibleNames
  Next wb
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function シートリスト取得()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.シートリスト取得"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    If sheetNameLists = "" Then
      sheetNameLists = tempSheet.Name
    Else
      sheetNameLists = sheetNameLists & vbNewLine & tempSheet.Name
    End If
  Next
  
  With Frm_Info
    .TextBox.Value = sheetNameLists
    .Show
  End With

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

