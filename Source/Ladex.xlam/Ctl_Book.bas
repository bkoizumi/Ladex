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
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Book.名前定義削除"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  
  For Each wb In Workbooks
    Workbooks(wb.Name).Activate
    Call Library.delVisibleNames
  Next wb
  
  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

'==================================================================================================
Function シートリスト取得()
  Dim tempSheet As Object
  Dim infoVal As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.シートリスト取得"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    If infoVal = "" Then
      infoVal = tempSheet.Name
    Else
      infoVal = infoVal & vbNewLine & tempSheet.Name
    End If
  Next
  
  With Frm_Info
    .TextBox.Value = infoVal
    .Show
  End With

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function
