Attribute VB_Name = "Ctl_TestCase"
Option Explicit


'==================================================================================================
Function セル範囲設定(line As Long, endLine As Long)

  Const funcName As String = "Ctl_TestCase.セル範囲設定"

  resultArea1 = "O" & line & ":S" & endLine
  resultArea2 = "U" & line & ":Y" & endLine
  resultArea3 = "AA" & line & ":AE" & endLine
  resultArea4 = "AG" & line & ":AK" & endLine
  resultArea5 = "AM" & line & ":AQ" & endLine



  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 再設定()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim categoryLine01 As Long, categoryLine02 As Long, categoryLine03 As Long
  
  Const funcName As String = "Ctl_TestCase.再設定"

  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  PrgP_Max = 4
  resetCellFlg = True
  runFlg = True
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  endLine = Range("A1").SpecialCells(xlLastCell).Row

  
  With Range("A" & startLine & ":AO" & Rows.count)
    '改行して全体を表示
    .WrapText = True
  
    'セルのロック
    .Locked = False
    .FormulaHidden = False
  End With
    
    

  Call Library.罫線_クリア(Range("A" & startLine & ":AQ" & endLine))
  
  Call Ctl_TestCase.セル範囲設定(startLine, endLine)
  
  Call Library.罫線_実線_格子(Range(resultArea1))
  Call Library.罫線_実線_格子(Range(resultArea2))
  Call Library.罫線_実線_格子(Range(resultArea3))
  Call Library.罫線_実線_格子(Range(resultArea4))
  Call Library.罫線_実線_格子(Range(resultArea5))
  
  
  For line = startLine To endLine - 1
    Call Ctl_TestCase.書式設定(line)
  Next
  Call Ctl_TestCase.書式設定(line)

  '連番の設定------------------------------------
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).ClearContents
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).NumberFormatLocal = "@"
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).HorizontalAlignment = xlCenter
  
  '書式設定--------------------------------------
  With Range("A" & startLine & ":A" & endLine)
    .HorizontalAlignment = xlLeft
    .ShrinkToFit = True
    .WrapText = False
  End With
  
  Range("H" & startLine & ":H" & endLine).HorizontalAlignment = xlCenter
  With Range("I" & startLine & ":I" & endLine & ",M" & startLine & ":M" & endLine)
    .HorizontalAlignment = xlLeft
    .ShrinkToFit = True
    .WrapText = False
  End With
  
  Range("O" & startLine & ":Q" & endLine & ",U" & startLine & ":W" & endLine & ",AA" & startLine & ":AC" & endLine & ",AG" & startLine & ":AI" & endLine & ",AM" & startLine & ":AO" & endLine).HorizontalAlignment = xlCenter
  Range("R" & startLine & ":S" & endLine & ",X" & startLine & ":Y" & endLine & ",AD" & startLine & ":AE" & endLine & ",AJ" & startLine & ":AK" & endLine & ",AP" & startLine & ":AQ" & endLine).HorizontalAlignment = xlLeft
  
  
  categoryLine01 = 1
  categoryLine02 = 1
  categoryLine03 = 1
  
  For line = startLine To endLine
    If Range("C" & line) <> "" Then
      Range("B" & line) = Format(WorksheetFunction.CountA(Range("C" & startLine & ":C" & line)), "00")
      categoryLine01 = 1
      categoryLine02 = 1
      categoryLine03 = 1
    End If
    
    If Range("E" & line) <> "" Then
      Range("D" & line) = Format(categoryLine01, "00")
      categoryLine01 = categoryLine01 + 1
      categoryLine02 = 1
      categoryLine03 = 1
    End If
    
    If Range("G" & line) <> "" Then
      Range("F" & line) = Format(categoryLine02, "00")
      categoryLine02 = categoryLine02 + 1
      categoryLine03 = 1
    End If
    
    If Range("L" & line) <> "" Then
      Range("K" & line) = Format(categoryLine03, "00")
      categoryLine03 = categoryLine03 + 1
    End If
    
  
    '条件付き書式設定----------------------------
    With Range("O" & line & ",U" & line & ",AA" & line & ",AG" & line & ",AM" & line)
      .FormatConditions.Delete
      
      .FormatConditions.Add Type:=xlTextString, String:="NG", TextOperator:=xlContains
      .FormatConditions(1).Font.Color = -16383844
      .FormatConditions(1).Font.Bold = True
      .FormatConditions(1).Font.TintAndShade = 0
      .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
      .FormatConditions(1).Interior.Color = 13551615
      .FormatConditions(1).Interior.TintAndShade = 0
    
      .FormatConditions.Add Type:=xlTextString, String:="要確認", TextOperator:=xlContains
      .FormatConditions(2).Font.Color = -16383844
      .FormatConditions(2).Font.Bold = True
      
      .FormatConditions.Add Type:=xlTextString, String:="既存バグ", TextOperator:=xlContains
      .FormatConditions(3).Font.Color = RGB(255, 255, 255)
      .FormatConditions(3).Font.Bold = True
      .FormatConditions(3).Font.TintAndShade = 0
      .FormatConditions(3).Interior.PatternColorIndex = xlAutomatic
      .FormatConditions(3).Interior.Color = RGB(83, 141, 213)
      .FormatConditions(3).Interior.TintAndShade = 0
    End With
    
    
  Next
  
  'テスト項目数の数式設定------------------------
  Range("Q8,W8,AC8,AI8,AO8").Formula = "=COUNTA(" & Range("G" & startLine & ":G" & endLine).Address & ")"
  'Range("Q9,W9,AC9,AI9,AO9").Formula = "=COUNTA(" & Range("J" & startLine & ":J" & endLine).Address & ")-COUNTIF(" & Range("M" & startLine & ":M" & endLine).Address & ",""-"")"
  Range("Q9,W9,AC9,AI9,AO9").Formula = "=COUNTA(" & Range("L" & startLine & ":L" & endLine).Address & ")"
  
  
  Range("Q10").Formula = "=COUNTIF(" & Range("O" & startLine & ":O" & endLine).Address & ",""OK"") + COUNTIF(" & Range("O" & startLine & ":O" & endLine).Address & ",""-"")"
  Range("Q11").Formula = "=COUNTIF(" & Range("O" & startLine & ":O" & endLine).Address & ",""NG"")"

  Range("W10").Formula = "=COUNTIF(" & Range("U" & startLine & ":U" & endLine).Address & ",""OK"") + COUNTIF(" & Range("U" & startLine & ":U" & endLine).Address & ",""-"")"
  Range("W11").Formula = "=COUNTIF(" & Range("U" & startLine & ":U" & endLine).Address & ",""NG"")"
  
  Range("AC10").Formula = "=COUNTIF(" & Range("AA" & startLine & ":AA" & endLine).Address & ",""OK"") + COUNTIF(" & Range("AA" & startLine & ":AA" & endLine).Address & ",""-"")"
  Range("AC11").Formula = "=COUNTIF(" & Range("AA" & startLine & ":AA" & endLine).Address & ",""NG"")"
  
  Range("AI10").Formula = "=COUNTIF(" & Range("AG" & startLine & ":AG" & endLine).Address & ",""OK"") + COUNTIF(" & Range("AG" & startLine & ":AG" & endLine).Address & ",""-"")"
  Range("AI11").Formula = "=COUNTIF(" & Range("AG" & startLine & ":AG" & endLine).Address & ",""NG"")"
    
  Range("AO10").Formula = "=COUNTIF(" & Range("AM" & startLine & ":AM" & endLine).Address & ",""OK"") + COUNTIF(" & Range("AM" & startLine & ":AM" & endLine).Address & ",""-"")"
  Range("AO11").Formula = "=COUNTIF(" & Range("AM" & startLine & ":AM" & endLine).Address & ",""NG"")"
    
    
  Call Ctl_TestCase.セル範囲設定(startLine, endLine)

  Call Library.罫線_実線_左(Range("B" & startLine & ":B" & endLine), , xlMedium)
  Call Library.罫線_実線_垂直(Range("B" & startLine & ":M" & endLine))
  Call Library.罫線_実線_右(Range("K" & startLine & ":M" & endLine), , xlMedium)
  Call Library.罫線_実線_下(Range("B" & startLine & ":M" & endLine), , xlMedium)


  Call Library.罫線_実線_囲み(Range("B" & startLine & ":M" & endLine), , xlMedium)

  Call Library.罫線_実線_囲み(Range(resultArea1), , xlMedium)
  Call Library.罫線_実線_囲み(Range(resultArea2), , xlMedium)
  Call Library.罫線_実線_囲み(Range(resultArea3), , xlMedium)
  Call Library.罫線_実線_囲み(Range(resultArea4), , xlMedium)
  Call Library.罫線_実線_囲み(Range(resultArea5), , xlMedium)
  
  
  'セルの高さ調整--------------------------------
  Rows(startLine & ":" & endLine).EntireRow.AutoFit
  
  For line = startLine To endLine
    If Rows(line & ":" & line).Height < 28 Then
      Rows(line & ":" & line).RowHeight = 28
    End If
  Next
  
  'セルの幅設定----------------------------------
  Range("O:O,U:U,AA:AA,AG:AG,AM:AM").ColumnWidth = 8.5
  Range("P:P,V:V,AB:AB,AH:AH,AN:AN").ColumnWidth = 15
  Range("Q:Q,W:W,AC:AC,AI:AI,AO:AO").ColumnWidth = 15
  Range("R:R,X:X,AD:AD,AJ:AJ,AP:AP").ColumnWidth = 33
  Range("S:S,Y:Y,AE:AE,AK:AK,AQ:AQ").ColumnWidth = 20
  
  
  '文字サイズ設定--------------------------------
  With Range("B" & startLine & ":BA" & endLine).Font
    .Name = "メイリオ"
    .Size = 9
  End With
  
  
  '文字の位置------------------------------------
  
  
  
  
  'データクリア----------------------------------
  'Range(resultArea1).ClearContents
  
  
  
   '入力規則設定----------------------------------
  With Range("H" & startLine & ":H" & endLine).Validation
    .Delete
    '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=SeleniumCmd"
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Library.covCells2String(Sh_Config.Range("D3:D" & Sh_Config.Cells(Rows.count, 4).End(xlUp).Row))
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .showError = True
  End With
  
 
  With Range( _
   "O" & startLine & ":O" & endLine & _
   ",U" & startLine & ":U" & endLine & _
   ",AA" & startLine & ":AA" & endLine & _
   ",AG" & startLine & ":AG" & endLine & _
   ",AM" & startLine & ":AM" & endLine).Validation
    
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Library.covCells2String(Sh_Config.Range("F3:F" & Sh_Config.Cells(Rows.count, 6).End(xlUp).Row))
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .showError = True
  End With
  
  
  
  'シートの保護----------------------------------
  'セルのロック
  'Range("M" & startLine & ":M" & endLine & ",S" & startLine & ":S" & endLine & ",Y" & startLine & ":Y" & endLine & ",AE" & startLine & ":AE" & endLine & ",AK" & startLine & ":AK" & endLine).Locked = True

  
  '処理終了--------------------------------------
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function 書式設定(targetLine As Long)

  Dim line As Long, endLine As Long
  Dim arryHeight(6) As Variant
  Dim setHeight As Long
  
  
  Const funcName As String = "Ctl_TestCase.書式設定"

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
'  If WorksheetFunction.CountA(Range("A" & targetLine & ":BB" & targetLine)) = 0 Then
'    Rows(targetLine & ":" & targetLine).Delete Shift:=xlUp
'    Exit Function
'  End If
  
  Call Ctl_TestCase.セル範囲設定(targetLine, targetLine)



  '中項目--------------------------------------
  If Range("C" & targetLine) <> "" Then
    Call Library.罫線_実線_上(Range("B" & targetLine & ":M" & targetLine), , xlMedium)
    
    Call Library.罫線_実線_上(Range(resultArea1), , xlMedium)
    Call Library.罫線_実線_上(Range(resultArea2), , xlMedium)
    Call Library.罫線_実線_上(Range(resultArea3), , xlMedium)
    Call Library.罫線_実線_上(Range(resultArea4), , xlMedium)
    Call Library.罫線_実線_上(Range(resultArea5), , xlMedium)
  Else
    Call Library.罫線_破線_上(Range(resultArea1))
    Call Library.罫線_破線_上(Range(resultArea2))
    Call Library.罫線_破線_上(Range(resultArea3))
    Call Library.罫線_破線_上(Range(resultArea4))
    Call Library.罫線_破線_上(Range(resultArea5))
  End If
  
  '小項目--------------------------------------
  If Range("E" & targetLine) <> "" Then
    If Range("D" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      Call Library.罫線_実線_上(Range("D" & targetLine & ":M" & targetLine))
        
      Call Library.罫線_実線_上(Range(resultArea1))
      Call Library.罫線_実線_上(Range(resultArea2))
      Call Library.罫線_実線_上(Range(resultArea3))
      Call Library.罫線_実線_上(Range(resultArea4))
      Call Library.罫線_実線_上(Range(resultArea5))
      
      
    End If
  End If
    
  'テスト項目--------------------------------
  If Range("G" & targetLine) <> "" Then
    If Range("F" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      'Call Library.罫線_破線_上(Range("B" & targetLine & ":B" & targetLine))
      Call Library.罫線_破線_上(Range("F" & targetLine & ":M" & targetLine))
    End If
  End If
  
  '入力データ---------------------------------
  If Range("H" & targetLine) <> "" Then
    If Range("H" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      'Call Library.罫線_破線_上(Range("B" & targetLine & ":B" & targetLine))
      
      Call Library.罫線_破線_上(Range("H" & targetLine & ":M" & targetLine))
    End If
  End If
    
  '確認事項---------------------------------
  If Range("I" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
    Call Library.罫線_破線_上(Range("H" & targetLine & ":M" & targetLine))
  End If




  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function シート追加()
  Dim sheetName As String
  Dim meg As String
  Const funcName As String = "Ctl_TestCase.シート追加"

  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  resetCellFlg = True
  runFlg = True
  '----------------------------------------------
  
  meg = "追加するシート名は？"
  
LBl_start:
  sheetName = InputBox(meg, "シート追加", "5.XXXXXX")
  
  If sheetName <> "" Then
    If Library.chkSheetExists(sheetName) = True Then
      meg = "存在するシート名は入力できません" & vbNewLine & "追加するシート名は？"
      
      GoTo LBl_start
    End If
    
    Sh_Copy.Copy After:=ActiveWorkbook.ActiveSheet
    ActiveSheet.Name = sheetName
    Call Library.startScript
  End If


  '処理終了--------------------------------------
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


