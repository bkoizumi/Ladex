Attribute VB_Name = "Ctl_Calendar"



'==================================================================================================
Function カレンダー削除()

  Const funcName As String = "Ctl_Calendar.カレンダー生成"
  
  
  Sh_WBS.Columns("W:XFD").Delete Shift:=xlToLeft
  
End Function


'==================================================================================================
Function カレンダー生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, endRowLine As Long
  Dim today As Date
  Dim HollydayName As String
  
  Const funcName As String = "Ctl_Calendar.カレンダー生成"

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
  resetCellFlg = True
  reCalflg = True
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Sh_WBS.Select
  Call Ctl_Calendar.カレンダー削除
  
  today = setVal("GUNT_START_DAY")
  line = Range(setVal("GUNT_START_COL_NM") & 1).Column
  
  
  Do While today <= setVal("GUNT_END_DAY")
    Cells(5, line) = today
    Call Library.罫線_破線_格子(Range(Cells(4, line), Cells(6, line)))
    If Format(today, "d") = 1 Or line = Range(setVal("GUNT_START_COL_NM") & 1).Column Then
      Cells(4, line) = today
      
      Call Library.罫線_実線_左(Range(Cells(4, line), Cells(6, line)))


    ElseIf DateSerial(Format(today, "yyyy"), Format(today, "m") + 1, 1) - 1 = today Or today = setVal("GUNT_END_DAY") Then
      '月末--------------------------------------
      Call Library.罫線_実線_右(Range(Cells(4, line), Cells(6, line)))
      
      Cells(4, line).Select
      Range(Selection, Selection.End(xlToLeft)).Merge
    Else

    End If
    
    '休日の設定==================================
    Call init.chkHollyday(today, HollydayName)
    Select Case HollydayName
      Case "Saturday"
        Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SaturdayColor")
        
      Case "Sunday"
        Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SundayColor")
      Case ""
      Case Else
        If HollydayName <> "会社指定休日" Then
          Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("SundayColor")
        Else
          Range(Cells(5, line), Cells(6, line)).Interior.Color = setVal("CompanyHolidayColor")
        End If
        '休日名をコメントに
        If TypeName(Cells(5, line).Comment) = "Nothing" Then
          Cells(5, line).AddComment HollydayName
        Else
          Cells(5, line).ClearComments
          Cells(5, line).AddComment HollydayName
        End If
        
        '期間中の休日リスト設定
    End Select
    
    '書式設定
    Cells(4, line).NumberFormatLocal = "m""月"""
    Cells(5, line).NumberFormatLocal = "d"
    
    line = line + 1
    today = today + 1
  Loop
  Rows("4:5").HorizontalAlignment = xlCenter
  Rows("4:5").ShrinkToFit = True
  Columns("W:XFD").ColumnWidth = 2
  
  
  
  
  Call Library.罫線_実線_囲み(Range(Cells(6, CInt(setVal("GUNT_START_COL"))), Cells(6, line - 1)))
        
  Call Library.罫線_実線_囲み(Range(Cells(4, CInt(setVal("GUNT_START_COL"))), Cells(4, line - 1)))
  Range(Cells(5, CInt(setVal("GUNT_START_COL"))), Cells(5, line - 1)).Select
  Call Library.setComment
    
  Range(Cells(3, line - 1), Cells(3, line - 1)).Select
'  Call 罫線.最終日
  
'  Range(setVal("calendarStartCol") & "1:" & setVal("calendarStartCol") & 6).Select
'  Call 罫線.二重線
'  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Copy
'  Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
'  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'  Range(Cells(3, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
'  Call 罫線.横線
'
'  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_LineInfo"))).End(xlUp).row
'  If endLine < 6 And Range(setVal("cell_TaskArea") & 6) = "" Then
'    endLine = 25
'  End If
'  Rows("6:" & endLine).Select
'  Selection.RowHeight = 20
'
'  Range("A6:B6").Select
'  Selection.Style = "数値"
'
'  Call 書式設定
'  Call 行書式コピー(6, endLine)
'
'  If ActiveSheet.Name = sheetMainName Then
'    Call init.名前定義
'  End If

  
  Application.Goto Reference:=Range("A1"), Scroll:=True
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
' * 行書式コピー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 行書式コピー(startLine As Long, endLine As Long)
  Dim line As Long
  Dim taskLevel As Long
  Dim taskLevelRange As Range
  Dim cell_LineInfo As Long
  
'  On Error GoTo catchError
  
  cell_LineInfo = 1
  'タスクが記載されている場合、タスクレベルを値としてコピー
  sheetMain.Calculate
  If Range(setVal("cell_TaskArea") & startLine) <> "" Then
    Range("B" & startLine & ":B" & endLine).Copy
    Range("B" & startLine & ":B" & endLine).PasteSpecial Paste:=xlPasteValues
  End If
  
  '書式のコピー＆ペースト
  Rows("4:4").Copy
  Rows(startLine & ":" & endLine).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  'タスクレベルの設定
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
    xlBetween, Formula1:="=担当者"
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
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function
