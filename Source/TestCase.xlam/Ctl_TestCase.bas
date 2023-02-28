Attribute VB_Name = "Ctl_TestCase"
Option Explicit


'==================================================================================================
Function �Z���͈͐ݒ�(line As Long, endLine As Long)

  Const funcName As String = "Ctl_TestCase.�Z���͈͐ݒ�"

  resultArea1 = "O" & line & ":S" & endLine
  resultArea2 = "U" & line & ":Y" & endLine
  resultArea3 = "AA" & line & ":AE" & endLine
  resultArea4 = "AG" & line & ":AK" & endLine
  resultArea5 = "AM" & line & ":AQ" & endLine



  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Đݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim categoryLine01 As Long, categoryLine02 As Long, categoryLine03 As Long
  
  Const funcName As String = "Ctl_TestCase.�Đݒ�"

  '�����J�n--------------------------------------
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
    '���s���đS�̂�\��
    .WrapText = True
  
    '�Z���̃��b�N
    .Locked = False
    .FormulaHidden = False
  End With
    
    

  Call Library.�r��_�N���A(Range("A" & startLine & ":AQ" & endLine))
  
  Call Ctl_TestCase.�Z���͈͐ݒ�(startLine, endLine)
  
  Call Library.�r��_����_�i�q(Range(resultArea1))
  Call Library.�r��_����_�i�q(Range(resultArea2))
  Call Library.�r��_����_�i�q(Range(resultArea3))
  Call Library.�r��_����_�i�q(Range(resultArea4))
  Call Library.�r��_����_�i�q(Range(resultArea5))
  
  
  For line = startLine To endLine - 1
    Call Ctl_TestCase.�����ݒ�(line)
  Next
  Call Ctl_TestCase.�����ݒ�(line)

  '�A�Ԃ̐ݒ�------------------------------------
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).ClearContents
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).NumberFormatLocal = "@"
  Range("B" & startLine & ":B" & endLine & ",D" & startLine & ":D" & endLine & ",F" & startLine & ":F" & endLine & ",K" & startLine & ":K" & endLine).HorizontalAlignment = xlCenter
  
  '�����ݒ�--------------------------------------
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
    
  
    '�����t�������ݒ�----------------------------
    With Range("O" & line & ",U" & line & ",AA" & line & ",AG" & line & ",AM" & line)
      .FormatConditions.Delete
      
      .FormatConditions.Add Type:=xlTextString, String:="NG", TextOperator:=xlContains
      .FormatConditions(1).Font.Color = -16383844
      .FormatConditions(1).Font.Bold = True
      .FormatConditions(1).Font.TintAndShade = 0
      .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
      .FormatConditions(1).Interior.Color = 13551615
      .FormatConditions(1).Interior.TintAndShade = 0
    
      .FormatConditions.Add Type:=xlTextString, String:="�v�m�F", TextOperator:=xlContains
      .FormatConditions(2).Font.Color = -16383844
      .FormatConditions(2).Font.Bold = True
      
      .FormatConditions.Add Type:=xlTextString, String:="�����o�O", TextOperator:=xlContains
      .FormatConditions(3).Font.Color = RGB(255, 255, 255)
      .FormatConditions(3).Font.Bold = True
      .FormatConditions(3).Font.TintAndShade = 0
      .FormatConditions(3).Interior.PatternColorIndex = xlAutomatic
      .FormatConditions(3).Interior.Color = RGB(83, 141, 213)
      .FormatConditions(3).Interior.TintAndShade = 0
    End With
    
    
  Next
  
  '�e�X�g���ڐ��̐����ݒ�------------------------
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
    
    
  Call Ctl_TestCase.�Z���͈͐ݒ�(startLine, endLine)

  Call Library.�r��_����_��(Range("B" & startLine & ":B" & endLine), , xlMedium)
  Call Library.�r��_����_����(Range("B" & startLine & ":M" & endLine))
  Call Library.�r��_����_�E(Range("K" & startLine & ":M" & endLine), , xlMedium)
  Call Library.�r��_����_��(Range("B" & startLine & ":M" & endLine), , xlMedium)


  Call Library.�r��_����_�͂�(Range("B" & startLine & ":M" & endLine), , xlMedium)

  Call Library.�r��_����_�͂�(Range(resultArea1), , xlMedium)
  Call Library.�r��_����_�͂�(Range(resultArea2), , xlMedium)
  Call Library.�r��_����_�͂�(Range(resultArea3), , xlMedium)
  Call Library.�r��_����_�͂�(Range(resultArea4), , xlMedium)
  Call Library.�r��_����_�͂�(Range(resultArea5), , xlMedium)
  
  
  '�Z���̍�������--------------------------------
  Rows(startLine & ":" & endLine).EntireRow.AutoFit
  
  For line = startLine To endLine
    If Rows(line & ":" & line).Height < 28 Then
      Rows(line & ":" & line).RowHeight = 28
    End If
  Next
  
  '�Z���̕��ݒ�----------------------------------
  Range("O:O,U:U,AA:AA,AG:AG,AM:AM").ColumnWidth = 8.5
  Range("P:P,V:V,AB:AB,AH:AH,AN:AN").ColumnWidth = 15
  Range("Q:Q,W:W,AC:AC,AI:AI,AO:AO").ColumnWidth = 15
  Range("R:R,X:X,AD:AD,AJ:AJ,AP:AP").ColumnWidth = 33
  Range("S:S,Y:Y,AE:AE,AK:AK,AQ:AQ").ColumnWidth = 20
  
  
  '�����T�C�Y�ݒ�--------------------------------
  With Range("B" & startLine & ":BA" & endLine).Font
    .Name = "���C���I"
    .Size = 9
  End With
  
  
  '�����̈ʒu------------------------------------
  
  
  
  
  '�f�[�^�N���A----------------------------------
  'Range(resultArea1).ClearContents
  
  
  
   '���͋K���ݒ�----------------------------------
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
  
  
  
  '�V�[�g�̕ی�----------------------------------
  '�Z���̃��b�N
  'Range("M" & startLine & ":M" & endLine & ",S" & startLine & ":S" & endLine & ",Y" & startLine & ":Y" & endLine & ",AE" & startLine & ":AE" & endLine & ",AK" & startLine & ":AK" & endLine).Locked = True

  
  '�����I��--------------------------------------
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function �����ݒ�(targetLine As Long)

  Dim line As Long, endLine As Long
  Dim arryHeight(6) As Variant
  Dim setHeight As Long
  
  
  Const funcName As String = "Ctl_TestCase.�����ݒ�"

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
'  If WorksheetFunction.CountA(Range("A" & targetLine & ":BB" & targetLine)) = 0 Then
'    Rows(targetLine & ":" & targetLine).Delete Shift:=xlUp
'    Exit Function
'  End If
  
  Call Ctl_TestCase.�Z���͈͐ݒ�(targetLine, targetLine)



  '������--------------------------------------
  If Range("C" & targetLine) <> "" Then
    Call Library.�r��_����_��(Range("B" & targetLine & ":M" & targetLine), , xlMedium)
    
    Call Library.�r��_����_��(Range(resultArea1), , xlMedium)
    Call Library.�r��_����_��(Range(resultArea2), , xlMedium)
    Call Library.�r��_����_��(Range(resultArea3), , xlMedium)
    Call Library.�r��_����_��(Range(resultArea4), , xlMedium)
    Call Library.�r��_����_��(Range(resultArea5), , xlMedium)
  Else
    Call Library.�r��_�j��_��(Range(resultArea1))
    Call Library.�r��_�j��_��(Range(resultArea2))
    Call Library.�r��_�j��_��(Range(resultArea3))
    Call Library.�r��_�j��_��(Range(resultArea4))
    Call Library.�r��_�j��_��(Range(resultArea5))
  End If
  
  '������--------------------------------------
  If Range("E" & targetLine) <> "" Then
    If Range("D" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      Call Library.�r��_����_��(Range("D" & targetLine & ":M" & targetLine))
        
      Call Library.�r��_����_��(Range(resultArea1))
      Call Library.�r��_����_��(Range(resultArea2))
      Call Library.�r��_����_��(Range(resultArea3))
      Call Library.�r��_����_��(Range(resultArea4))
      Call Library.�r��_����_��(Range(resultArea5))
      
      
    End If
  End If
    
  '�e�X�g����--------------------------------
  If Range("G" & targetLine) <> "" Then
    If Range("F" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      'Call Library.�r��_�j��_��(Range("B" & targetLine & ":B" & targetLine))
      Call Library.�r��_�j��_��(Range("F" & targetLine & ":M" & targetLine))
    End If
  End If
  
  '���̓f�[�^---------------------------------
  If Range("H" & targetLine) <> "" Then
    If Range("H" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
      'Call Library.�r��_�j��_��(Range("B" & targetLine & ":B" & targetLine))
      
      Call Library.�r��_�j��_��(Range("H" & targetLine & ":M" & targetLine))
    End If
  End If
    
  '�m�F����---------------------------------
  If Range("I" & targetLine).Borders(xlEdgeTop).LineStyle <> 1 Then
    Call Library.�r��_�j��_��(Range("H" & targetLine & ":M" & targetLine))
  End If




  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function �V�[�g�ǉ�()
  Dim sheetName As String
  Dim meg As String
  Const funcName As String = "Ctl_TestCase.�V�[�g�ǉ�"

  '�����J�n--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  resetCellFlg = True
  runFlg = True
  '----------------------------------------------
  
  meg = "�ǉ�����V�[�g���́H"
  
LBl_start:
  sheetName = InputBox(meg, "�V�[�g�ǉ�", "5.XXXXXX")
  
  If sheetName <> "" Then
    If Library.chkSheetExists(sheetName) = True Then
      meg = "���݂���V�[�g���͓��͂ł��܂���" & vbNewLine & "�ǉ�����V�[�g���́H"
      
      GoTo LBl_start
    End If
    
    Sh_Copy.Copy After:=ActiveWorkbook.ActiveSheet
    ActiveSheet.Name = sheetName
    Call Library.startScript
  End If


  '�����I��--------------------------------------
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


