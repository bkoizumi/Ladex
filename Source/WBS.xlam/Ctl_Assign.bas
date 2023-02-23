Attribute VB_Name = "Ctl_Assign"
Option Explicit

'==================================================================================================
Function 担当者リスト表示()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim i As Integer
  Const funcName As String = "Ctl_Assign.担当者リスト表示"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  

  '担当者読み込み--------------------------------
  Erase lstAssign
  endLine = sh_Option.Cells(Rows.count, 11).End(xlUp).Row
  i = 0
  On Error Resume Next
  For line = 4 To endLine
    If sh_Option.Range("L" & line) <> "" Then
      ReDim Preserve lstAssign(i)
      
      lstAssign(i) = sh_Option.Range("K" & line).Text
      Call Library.showDebugForm("担当者：", sh_Option.Range("K" & line).Text, "debug")
      
      i = i + 1
    End If
  Next
  
  




  '処理終了--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
