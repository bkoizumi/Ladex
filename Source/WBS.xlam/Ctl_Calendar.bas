Attribute VB_Name = "Ctl_Calendar"



'==================================================================================================
Function �J�����_�[�폜()

  Const funcName As String = "Ctl_Calendar.�J�����_�[����"
  
  
  Sh_WBS.Columns("W:XFD").Delete Shift:=xlToLeft
  
End Function


'==================================================================================================
Function �J�����_�[����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, endRowLine As Long
  Dim today As Date
  Dim HollydayName As String
  
  Const funcName As String = "Ctl_Calendar.�J�����_�[����"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  resetCellFlg = True
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Sh_WBS.Select
  Call Ctl_Calendar.�J�����_�[�폜
  
  today = setVal("GUNT_START_DAY")
  line = Range(setVal("GUNT_START_COL_NM") & 1).Column
  
  
  Do While today <= setVal("GUNT_END_DAY")
    Cells(5, line) = today
    Call Library.�r��_�j��_�i�q(Range(Cells(4, line), Cells(6, line)))
    If Format(today, "d") = 1 Or line = Range(setVal("GUNT_START_COL_NM") & 1).Column Then
      Cells(4, line) = today
      
      Call Library.�r��_����_��(Range(Cells(4, line), Cells(6, line)))


    ElseIf DateSerial(Format(today, "yyyy"), Format(today, "m") + 1, 1) - 1 = today Or today = setVal("GUNT_END_DAY") Then
      '����--------------------------------------
      Call Library.�r��_����_�E(Range(Cells(4, line), Cells(6, line)))
      
      Cells(4, line).Select
      Range(Selection, Selection.End(xlToLeft)).Merge
    Else

    End If
    
    '�x���̐ݒ�==================================
    Call init.chkHollyday(today, HollydayName)
    Select Case HollydayName
      Case "Saturday"
        Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SaturdayColor")
        
      Case "Sunday"
        Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SundayColor")
      Case ""
      Case Else
        If HollydayName <> "��Ўw��x��" Then
          Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SundayColor")
        Else
          Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("CompanyHolidayColor")
        End If
        '�x�������R�����g��
        If TypeName(Cells(5, line).Comment) = "Nothing" Then
          Cells(5, line).AddComment HollydayName
        Else
          Cells(5, line).ClearComments
          Cells(5, line).AddComment HollydayName
        End If
        
        '���Ԓ��̋x�����X�g�ݒ�
    End Select
    
    '�����ݒ�
    Cells(4, line).NumberFormatLocal = "m""��"""
    Cells(5, line).NumberFormatLocal = "d"
    
    line = line + 1
    today = today + 1
  Loop
  Rows("4:5").HorizontalAlignment = xlCenter
  Rows("4:5").ShrinkToFit = True
  Columns("W:XFD").ColumnWidth = 2
  
  
  
  
  Call Library.�r��_����_�͂�(Range(Cells(6, CInt(setVal("GUNT_START_COL"))), Cells(6, line - 1)))
        
  Call Library.�r��_����_�͂�(Range(Cells(4, CInt(setVal("GUNT_START_COL"))), Cells(4, line - 1)))
  Range(Cells(5, CInt(setVal("GUNT_START_COL"))), Cells(5, line - 1)).Select
  Call Library.setComment
    
  Range(Cells(3, line - 1), Cells(3, line - 1)).Select
'  Call �r��.�ŏI��
  
'  Range(setVal("calendarStartCol") & "1:" & setVal("calendarStartCol") & 6).Select
'  Call �r��.��d��
'  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Copy
'  Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
'  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'  Range(Cells(3, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
'  Call �r��.����
'
'  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_LineInfo"))).End(xlUp).row
'  If endLine < 6 And Range(setVal("cell_TaskArea") & 6) = "" Then
'    endLine = 25
'  End If
'  Rows("6:" & endLine).Select
'  Selection.RowHeight = 20
'
'  Range("A6:B6").Select
'  Selection.Style = "���l"
'
'  Call �����ݒ�
'  Call �s�����R�s�[(6, endLine)
'
'  If ActiveSheet.Name = sheetMainName Then
'    Call init.���O��`
'  End If

  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  '�����I��--------------------------------------
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

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * �s�����R�s�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �s�����R�s�[(startLine As Long, endLine As Long)
  Dim line As Long
  Dim taskLevel As Long
  Dim taskLevelRange As Range
  Dim cell_LineInfo As Long
  
'  On Error GoTo catchError
  
  cell_LineInfo = 1
  '�^�X�N���L�ڂ���Ă���ꍇ�A�^�X�N���x����l�Ƃ��ăR�s�[
  sheetMain.Calculate
  If Range(setVal("cell_TaskArea") & startLine) <> "" Then
    Range("B" & startLine & ":B" & endLine).Copy
    Range("B" & startLine & ":B" & endLine).PasteSpecial Paste:=xlPasteValues
  End If
  
  '�����̃R�s�[���y�[�X�g
  Rows("4:4").Copy
  Rows(startLine & ":" & endLine).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  '�^�X�N���x���̐ݒ�
  If ActiveSheet.Name = sheetMainName Then
    For line = 6 To endLine
      If Range(setVal("cell_TaskArea") & line) <> "" Then
        taskLevel = Range(setVal("cell_LevelInfo") & line) - 1
        If taskLevel > 0 Then
          Range(setVal("cell_TaskArea") & line).InsertIndent taskLevel
        End If
      End If
      
      If Range(setVal("cell_Info") & line) <> setVal("TaskInfoStr_Multi") Then
        Range("A" & line) = cell_LineInfo
        cell_LineInfo = cell_LineInfo + 1
      Else
        Range("A" & line) = Range("A" & line - 1)
      End If
      
      Range(setVal("cell_LineInfo") & line).FormulaR1C1 = "=ROW()-5"
      Set taskLevelRange = Range(setVal("cell_TaskArea") & line)
      Range(setVal("cell_LevelInfo") & line).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
      Set taskLevelRange = Nothing
    Next
  End If
  
  With Range(setVal("cell_Assign") & startLine & ":" & setVal("cell_Assign") & endLine).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="=�S����"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .showError = False
  End With

  With Range(setVal("cell_TaskArea") & startLine & ":" & setVal("cell_TaskArea") & endLine).Validation
    .Delete
    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
    :=xlBetween
    .IgnoreBlank = True
    .InCellDropdown = False
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOn
    .ShowInput = True
    .showError = True
  End With
  

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function
