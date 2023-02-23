Attribute VB_Name = "Ctl_Sumally"
Option Explicit


'**************************************************************************************************
' * 月別進捗表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setMonthly()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim suLine As Long
  
  Const funcName As String = "Ctl_Sumally.setSumally"

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
  

  If Library.chkSheetExists("Sumally") = False Then
    ThisWorkbook.Worksheets("Sumally").Copy After:=Sh_WBS
  End If

  Set sh_Sumally = ActiveWorkbook.Worksheets("Sumally")
  Call Library.delSheetData(sh_Sumally, 4)

  endLine = Sh_WBS.Cells(Rows.count, 1).End(xlUp).Row
  suLine = 4
  
  sh_Sumally.Range("H1") = "基準日：" & Format(setVal("KIJUNBI"), "M/D")
  For line = setVal("TASK_START_ROW") To endLine
    If Sh_WBS.Range("B" & line) = 1 Then
      sh_Sumally.Range("A" & suLine) = suLine - 3
      sh_Sumally.Range("B" & suLine) = Sh_WBS.Range("C" & line)
      sh_Sumally.Range("C" & suLine) = Sh_WBS.Cells(line, Int(setVal("PROG_COL")))
      sh_Sumally.Range("D" & suLine) = Sh_WBS.Cells(line, Int(setVal("PLAN_START_COL")))
      sh_Sumally.Range("E" & suLine) = Sh_WBS.Cells(line, Int(setVal("PLAN_END_COL")))
      
      sh_Sumally.Range("F" & suLine) = Sh_WBS.Cells(line, Int(setVal("ACT_START_COL")))
      sh_Sumally.Range("G" & suLine) = Sh_WBS.Cells(line, Int(setVal("ACT_END_COL")))
      
      sh_Sumally.Range("H" & suLine) = Sh_WBS.Range("U" & line)
      
      
      suLine = suLine + 1
    End If
  Next
  
  '書式設定--------------------------------------
  sh_Sumally.Columns("A:A").HorizontalAlignment = xlCenter
  sh_Sumally.Columns("D:G").NumberFormatLocal = "m/d"
  sh_Sumally.Columns("H:H").NumberFormatLocal = "0.00"
  
  sh_Sumally.Rows("4:" & suLine).VerticalAlignment = xlCenter
  


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
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function

