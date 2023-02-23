Attribute VB_Name = "Ctl_Task"
Option Explicit

'**************************************************************************************************
' * タスク操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 基準日設定()
  Dim nowTime As Integer
  Dim today As Date
  
  Const funcName As String = "Ctl_Task.基準日設定"
  Call Library.showDebugForm(funcName, , "start1")
  
  today = Format(Date, "yyyy/mm/dd")
  nowTime = Format(Now(), "h")

  Select Case nowTime
    Case Is <= 14
      Sh_PARAM.Range("B13") = Format(DateAdd("d", -1, Date), "yyyy/mm/dd")
      Sh_PARAM.Range("B14") = "FALSE"
      Sh_PARAM.Range("B15") = "TRUE"
    Case Else
      Sh_PARAM.Range("B13") = Format(Date, "yyyy/mm/dd")
      Sh_PARAM.Range("B14") = "TRUE"
      Sh_PARAM.Range("B15") = "FALSE"
  End Select

End Function


'==================================================================================================
Function タスクチェック()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim staffName As String
  Dim targetCell As Range

  Const funcName As String = "Ctl_Task.タスクチェック"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  resetCellFlg = True
  
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  Call Ctl_Task.基準日設定
  Call Ctl_Style.書式設定
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  endLine = Range("A1").SpecialCells(xlLastCell).Row
  
  endColLine = Cells(5, Columns.count).End(xlToLeft).Column
  
  Cells.Font.Name = "メイリオ"
  Cells.Font.Size = 9
  
  Range(Cells(6, 23), Cells(6, endColLine)).Copy
  Range(Cells(7, 23), Cells(endLine, endColLine)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

  
  Call Library.罫線_破線_水平(Range("A7:V" & endLine))
  Call Library.罫線_実線_囲み(Range("C7:H" & endLine))
  
  
  Call Library.罫線_実線_囲み(Range("C7:H" & endLine))
  Set targetCell = Range("A7:A" & endLine & ",B7:B" & endLine & ",I7:J" & endLine & ",K7:L" & endLine & ",M7:N" & endLine & ",O7:P" & endLine & ",Q7:R" & endLine & ",S7:T" & endLine & ",V7:V" & endLine)
  Call Library.罫線_実線_囲み(targetCell)
  Call Library.罫線_破線_水平(targetCell)
  
  Set targetCell = Range("I7:J" & endLine & ",K7:L" & endLine & ",M7:N" & endLine & ",O7:P" & endLine & ",Q7:R" & endLine & ",S7:T" & endLine)
  Call Library.罫線_破線_垂直(targetCell)
  Set targetCell = Nothing
  
  
  Range("A" & startLine & ":A" & endLine).FormulaR1C1 = "=ROW()-6"
'  Range("B" & startLine & ":B" & endLine).FormulaR1C1 = "=IF(RC[1]<>"""",1,IF(RC[2]<>"""",2,IF(RC[3]<>"""",3,IF(RC[4]<>"""",4,IF(RC[5]<>"""",5,IF(RC[6]<>"""",6,0)" & Chr(10) & ")))))"
'  ActiveSheet.Calculate
  
  
  For line = startLine To endLine
    If Range("B" & line) <> Range("C" & line).IndentLevel + 1 Then
      Range("C" & line).IndentLevel = Range("B" & line) - 1
    End If
    
    
    
    'タスクレベル1なら実線-----------------------
    Set targetCell = Range(Cells(line, 1), Cells(line, endColLine))
    
    If Range("B" & line) = 1 Then
      Call Library.罫線_実線_上(targetCell)
    Else
      Call Library.罫線_破線_上(targetCell)
    End If
    Set targetCell = Nothing
    
    '工程タスク調査------------------------------
    If Range("B" & line) < Range("B" & line + 1) Then
      If Range("L" & line) = "" Then Range("L" & line) = "工程"
    End If
  
  
    '担当者の色付け------------------------------
    staffName = ""
    If Range("L" & line) <> "" Then
      staffName = Range("L" & line)
    ElseIf Range("K" & line) <> "" Then
      staffName = Range("K" & line)
    End If
    
    Call Library.showDebugForm("staffName", staffName & "<>" & setAssign(staffName), "debug")
    If staffName <> "" Then
      If setAssign(staffName) <> "" Then
        Range("B" & line).Interior.Color = setAssign(staffName)
      End If
    Else
      If Range("B" & line).Style = "Normal" Then
        Range("B" & line).Interior.ColorIndex = 0
      End If
    End If
    
'    If staffName = "TBA" Or staffName = "TBC" Or staffName = "TBD" Then
'      Range("K" & line & ":L" & line).Font.Bold = True
'    End If

    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "タスクチェック")
  Next
  
  Set targetCell = Range(Cells(endLine, 1), Cells(endLine, endColLine))
  Call Library.罫線_破線_下(targetCell)
  Set targetCell = Nothing
  
  
  '担当者----------------------------------------
  Call Ctl_Assign.担当者リスト表示

  With Range("K" & startLine & ":L" & endLine).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(lstAssign, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOn
    .ShowInput = True
    .showError = False
  End With
  
  
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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

'**************************************************************************************************
' * タスク移動などの操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function タスク追加()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク追加"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  
  slctStartLine = Selection.Row
  If Selection.Rows.count > 1 Then
    slctEndLine = Selection.Row + Selection.Rows.count - 1
  Else
    slctEndLine = Selection.Row
  End If
  
  'Rows(slctStartLine & ":" & slctEndLine).Select
  Rows(slctStartLine & ":" & slctEndLine).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  
  
  Range("A" & slctStartLine & ":A" & slctEndLine).FormulaR1C1 = "=ROW()-6"
  Range("B" & slctStartLine & ":B" & slctEndLine).FormulaR1C1 = "=IF(RC[1]<>"""",1,IF(RC[2]<>"""",2,IF(RC[3]<>"""",3,IF(RC[4]<>"""",4,IF(RC[5]<>"""",5,IF(RC[6]<>"""",6,"""")" & Chr(10) & ")))))"
  
  Call Ctl_Task.タスクチェック
  
  Range("C" & slctStartLine).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function タスク削除()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク削除"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  
  slctStartLine = Selection.Row
  If Selection.Rows.count > 1 Then
    slctEndLine = Selection.Row + Selection.Rows.count - 1
  Else
    slctEndLine = Selection.Row
  End If
  
  'Rows(slctStartLine & ":" & slctEndLine).Select
  Rows(slctStartLine & ":" & slctEndLine).Delete Shift:=xlUp
  Range("C" & slctStartLine).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function タスク移動_上()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク移動_上"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  Set selectedCells = Selection
  
  slctStartLine = Selection.Row
  If Selection.Rows.count > 1 Then
    slctEndLine = Selection.Row + Selection.Rows.count - 1
  Else
    slctEndLine = Selection.Row
  End If
  
  If slctStartLine - 1 >= 7 Then
    Rows(slctStartLine & ":" & slctEndLine).Cut
    Rows(slctStartLine - 1 & ":" & slctEndLine - 1).Insert Shift:=xlDown
  End If
  
  selectedCells.Select
  Set selectedCells = Nothing
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function タスク移動_下()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク移動_下"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Set selectedCells = Selection
  
  slctStartLine = Selection.Row
  If Selection.Rows.count > 1 Then
    slctEndLine = Selection.Row + Selection.Rows.count - 1
  Else
    slctEndLine = Selection.Row
  End If
  
  'Rows(slctStartLine & ":" & slctEndLine).Select
  Rows(slctStartLine & ":" & slctEndLine).Cut
  Rows(slctEndLine + 2 & ":" & slctEndLine + 2).Insert Shift:=xlDown
  
  
  selectedCells.Select
  Set selectedCells = Nothing
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function タスク移動_左()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク移動_左"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For line = Selection(1).Row To Selection(Selection.count).Row
    Range("C" & line).IndentLevel = Range("C" & line).IndentLevel + 1
    Range("B" & line) = Range("C" & line).IndentLevel + 1
  Next



  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function タスク移動_右()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctStartLine As Long, slctEndLine As Long
  Dim selectedCells As Range, targetCell As Range
  
  Const funcName As String = "Ctl_Task.タスク移動_右"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  'Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  
  For line = Selection(1).Row To Selection(Selection.count).Row
    Range("C" & line).IndentLevel = Range("C" & line).IndentLevel - 1
    Range("B" & line) = Range("C" & line).IndentLevel + 1
  Next

  Selection.Offset(, 1).Select
  
  
  '処理終了--------------------------------------
  'Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 進捗率設定(progress As Long)
  Dim line As Long
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Task.進捗率設定"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctCells In Selection
    line = slctCells.Row
    Range("J" & line) = progress
    
    DoEvents
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function タスクのリンク設定()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
  
    
  Const funcName As String = "Ctl_Task.タスクのリンク設定"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 3
  Else
'    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  PBarCnt = 1
  PrgP_Max = 3
  
  Call Ctl_ProgressBar.showStart
  setVal("debugMode") = "develop"
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  
  '----------------------------------------------

  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    If Range("K" & targetCell.Row) = "工程" Or Range("L" & targetCell.Row) = "工程" Then
    
    ElseIf oldLine = 0 Then
      Range("N" & targetCell.Row) = Library.CalAddDay(Range("M" & targetCell.Row), Range("S" & targetCell.Row), , "end")
      oldLine = targetCell.Row
    Else
      If Format(Range("N" & oldLine), "h") = 0 Or Format(Range("N" & oldLine), "h") > 14 Then
        
        '先行タスクの終了日+1を開始日に設定
        newStartDay = Format(Range("N" & oldLine), "yyyy/mm/dd 09:00:00")
        'newStartDay = DateAdd("d", 1, newStartDay)
        
        Range("M" & targetCell.Row) = Library.CalAddDay(newStartDay, 1, "day")
        'Range("M" & targetCell.Row) = Format(Range("M" & targetCell.Row), "yyyy/mm/dd 09:00:00")
      Else
        Range("M" & targetCell.Row) = Range("N" & oldLine)
      End If
      
      '終了日を再設定
      Range("N" & targetCell.Row) = Library.CalAddDay(Range("M" & targetCell.Row), Range("S" & targetCell.Row), , "end")
      oldLine = targetCell.Row
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, PBarCnt, selectedCells.count, "タスクのリンク設定")
    PBarCnt = PBarCnt + 1
  Next

  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
'    Application.Goto Reference:=Range("A1"), Scroll:=True
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

'==================================================================================================
Function タスクのリンク解除()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
    
  Const funcName As String = "Ctl_Task.タスクのリンク解除"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
   
  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Range(setVal("cell_Task") & targetCell.Row) = ""
  Next


  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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


'==================================================================================================
Function タスクにスクロール()
  Dim line As Long, activeCellRowLine As Long, activeCellColLine As Long
  Dim targetColumn As Integer
  
  Const funcName As String = "Ctl_Task.タスクにスクロール"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  activeCellRowLine = ActiveCell.Row
  activeCellColLine = ActiveCell.Column
  
  If Range("M" & activeCellRowLine).Text <> "" Then
    targetColumn = Library.getColumnNo(WBS_Option.日付セル検索(Range("M" & activeCellRowLine).Text))
  Else
    targetColumn = Library.getColumnNo(WBS_Option.日付セル検索(Worksheets("PARAM").Range("B13")))
  End If
  ActiveWindow.ScrollColumn = targetColumn - 10
  
  
  Cells(activeCellRowLine, targetColumn).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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


'==================================================================================================
Function 進捗コピー()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "Ctl_Task.進捗コピー"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  'Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
 
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  Range("J7:J" & endLine).Copy
  Range("I7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  
  '処理終了--------------------------------------
  'Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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
