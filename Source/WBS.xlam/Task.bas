Attribute VB_Name = "Task"
'**************************************************************************************************
' * �^�X�N�����o
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�����o(taskList As Collection)
  Dim line As Long, endLine As Long, count As Long

'  On Error GoTo catchError

  Call init.setting
  Set taskList = New Collection
  count = 1
  
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).Row
  count = count + 1
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_DataExtract") & line) <> "" Then
      With taskList
        .Add item:=setSheet.Range(setVal("cell_DataExtract") & line).Value, Key:=str(count)
      End With
      count = count + 1
    End If
  Next
  Exit Function
  
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �S���Ғ��o
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �S���Ғ��o(memberList As Collection)
  Dim line As Long, endLine As Long, count As Long
  Dim assignor As String
  
  
'  On Error GoTo catchError

  Call init.setting
  Set memberList = New Collection
  count = 1
  
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).Row
  With memberList
    .Add item:="�H��", Key:=str(count)
  End With
  count = count + 1
  NoAssignorFlg = False
  
  For line = 6 To endLine
    assignor = mainSheet.Range(setVal("cell_Assign") & line).Value
    If assignor <> "" Then
        If isCollection(memberList, assignor) = False Then
          With memberList
            .Add item:=assignor, Key:=str(count)
          End With
          count = count + 1
        End If
    
    ElseIf assignor = "" And NoAssignorFlg = False Then
      With memberList
        .Add item:="�����蓖��", Key:=str(count)
      End With
      count = count + 1
      NoAssignorFlg = True
    End If
  Next





'  For line = 6 To endLine
'    If mainSheet.Range(setVal("cell_Assign") & line).Value <> "" Then
'      For Each assignName In Split(mainSheet.Range(setVal("cell_Assign") & line).Value, ",")
'        assignor = assignName
'        If assignor <> "" And isCollection(memberList, assignor) = False Then
'          With memberList
'            .Add item:=assignor, Key:=str(count)
'          End With
'          count = count + 1
'        End If
'      Next
'    End If
'  Next
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function



Function isCollection(col As Collection, query) As Boolean
  Dim item
  
  For Each item In col
    If item = query Then
      isCollection = True
      Exit Function
    End If
  Next
  isCollection = False
End Function


'**************************************************************************************************
' * �S���҃t�B���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �S���҃t�B���^�[(filterName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  On Error GoTo catchError

  Unload FilterForm
  Call Library.startScript
  Call ProgressBar.showStart
  Call init.setting
  
  mainSheet.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("�S���҃t�B���^�[", line, endLine, "")
    
    If Range(setVal("cell_Assign") & line).Text = filterName Or Range(setVal("cell_Assign") & line).Text = filterName Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
  
'**************************************************************************************************
' * �^�X�N���t�B���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N���t�B���^�[(filterNames As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  On Error GoTo catchError
  Call Library.showDebugForm("�^�X�N���t�B���^�[", "�J�n")

  Unload FilterForm
  Call Library.startScript
  Call init.setting
  
  mainSheet.Select
  
  '��\���s��S�ĕ\��
  Cells.EntireRow.Hidden = False
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  For line = 6 To endLine
    DoEvents
    For Each filterName In Split(filterNames, "<>")
      DoEvents
      
      If Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_Multi") Then
        Rows(line & ":" & line).EntireRow.Hidden = True
      ElseIf Range(setVal("cell_TaskArea") & line) Like "*" & filterName & "*" Then
        Rows(line & ":" & line).EntireRow.Hidden = False
        Exit For
      Else
        Rows(line & ":" & line).EntireRow.Hidden = True
      End If
    Next
  Next
  
  
  Call Library.endScript
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �i���R�s�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �i���R�s�[()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Task.�i���R�s�["

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
 
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  Range("J7:J" & endLine).Copy
  Range("I7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������=====================================================================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'**************************************************************************************************
' * �^�X�N�̑}��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�̑}��()
  Dim taskLevelRange As Range
'  On Error GoTo catchError
  

  Rows("4:4").Copy
  Rows(Selection.Row & ":" & Selection.Row).Insert Shift:=xlDown
  Range("A" & Selection.Row).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Range(setVal("cell_Info") & Selection.Row & ":XFD" & Selection.Row).ClearContents
  Range(setVal("cell_Info") & Selection.Row & ":XFD" & Selection.Row).ClearComments
  
  Range("A" & Selection.Row) = Range("A" & Selection.Row - 1) + 1
  
  
  Set taskLevelRange = Range(setVal("cell_TaskArea") & Selection.Row)
  Range(setVal("cell_LevelInfo") & Selection.Row).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
  Set taskLevelRange = Nothing

  
  Range(setVal("cell_LineInfo") & Selection.Row).FormulaR1C1 = "=ROW()-5"
 
  Call WBS_Option.�s�ԍ��Đݒ�

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �^�X�N�̍폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�̍폜()
  Dim selectedCells As Range

'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  mainSheet.Select


  Rows(Selection(1).Row & ":" & Selection(Selection.count).Row).Delete Shift:=xlUp
  Call WBS_Option.�s�ԍ��Đݒ�

  Call Library.endScript

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function




'**************************************************************************************************
' * �^�X�N���擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getTaskName(Optional line As Long = 0, Optional retFlg As String = "value")
  Dim endLine As Long, colLine As Long, endColLine As Long
  Dim TaskRange As Range
  
  Const funcName As String = "Task.getTaskName"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  If line = 0 Then
    line = ActiveCell.Row
  End If
  
  For colLine = 3 To 8
    If Cells(line, colLine) <> "" Then
      Set TaskRange = Cells(line, colLine)
      Exit For
    End If
  Next
  Call Library.showDebugForm("TaskName", TaskRange.Text, "debug")
  
  If retFlg = "value" Then
    getTaskName = TaskRange.Text
  Else
    getTaskName = TaskRange.Address
  End If
  
  Set TaskRange = Nothing
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
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
