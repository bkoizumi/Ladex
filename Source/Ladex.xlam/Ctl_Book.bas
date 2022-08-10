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
  Const funcName As String = "Ctl_Book.名前定義削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
'  For Each wb In Workbooks
'    Workbooks(wb.Name).Activate
'    Call Library.delVisibleNames
'  Next wb
  
  Call Library.delVisibleNames
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
     Call Library.showDebugForm(funcName, , "start1")
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
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 印刷範囲の点線を非表示()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.印刷範囲の点線を非表示"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = False
  Next
  

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 印刷範囲の点線を表示()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.印刷範囲の点線を表示"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = True
  Next
  

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
