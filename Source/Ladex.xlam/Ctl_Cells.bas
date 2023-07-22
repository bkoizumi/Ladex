Attribute VB_Name = "Ctl_Cells"
Option Explicit

'**************************************************************************************************
' * �Z������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �Z����������_��()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  Dim colName As String
  
  Const funcName As String = "Ctl_Cells.�Z����������_��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  Columns(Library.getColumnName(startColLine) & ":" & Library.getColumnName(endColLine)).EntireColumn.AutoFit
  
  '�ő�l�m�F------------------------------------
  For colLine = startLine To endColLine
    colName = Library.getColumnName(colLine)
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, colLine, endColLine, "")
    
    If Columns(colLine).Hidden = False Then
      '�񕝂̎�������
      Columns(colLine).EntireColumn.AutoFit
    
      If Columns(colLine).ColumnWidth > maxColumnWidth Then
        Columns(colLine).ColumnWidth = maxColumnWidth
      End If
    End If
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z����������_����()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z����������_����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
    
    If Rows(line).Hidden = False Then
      '�����̎�������
      Rows(line).EntireRow.AutoFit
      
      If Rows(line).Height > maxRowHeight Then
        Rows(line).rowHeight = maxRowHeight
      End If
    End If
  Next
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z����������_����()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z����������_����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Max = 2
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Ctl_Cells.�Z����������_����
  Call Ctl_Cells.�Z����������_��
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z���Œ�ݒ�_��()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z���Œ�ݒ�_��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  Columns(Library.getColumnName(startColLine) & ":" & Library.getColumnName(endColLine)).EntireColumn.AutoFit
  
  '�ő�l�m�F------------------------------------
  For colLine = startLine To endColLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, colLine, endColLine, "")
    
    If Columns(colLine).Hidden = False Then
      Columns(colLine).ColumnWidth = dicVal("ColumnWidth")
    End If
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z���Œ�ݒ�_����()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z���Œ�ݒ�_����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
    If Rows(line).Hidden = False Then
      Rows(line).rowHeight = dicVal("rowHeight")
    End If
  Next
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z���Œ�ݒ�_����()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z���Œ�ݒ�_����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  PrgP_Max = 2
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  Ctl_Cells.�Z���Œ�ݒ�_����
  Ctl_Cells.�Z���Œ�ݒ�_��
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z���Œ�ݒ�_�����w��(heightVal As Integer)
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.�Z���Œ�ݒ�_����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  '�I��͈͂�����ꍇ----------------------------
  If Selection.Rows.count > 1 Then
    startLine = Selection(1).Row
    endLine = Selection(Selection.count).Row
  
    For line = startLine To endLine
      Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
      
      If Rows(line).Hidden = False Then
        Rows(line).rowHeight = heightVal
      End If
    Next
  Else
    Cells.rowHeight = heightVal
  End If
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'**************************************************************************************************
' * �Z���ҏW
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �폜_�O��̃X�y�[�X()
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.�폜_�O��̃X�y�[�X"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    slctCells.Value = Trim(slctCells.Text)
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �폜_�S�X�y�[�X()
  Dim slctCells As Range
  Dim resVal As String
  
  Const funcName As String = "Ctl_Cells.�폜_�S�X�y�[�X"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    resVal = slctCells.Text

    If resVal <> "" Then
      resVal = Replace(resVal, " ", "")
      resVal = Replace(resVal, "�@", "")
      slctCells.Value = resVal
      DoEvents
    End If
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ǉ�_�����ɒ����_()
  Dim line As Long, endLine As Long
  Dim Reg As Object
  Dim slctCells
  
  Const funcName As String = "Ctl_Cells.�ǉ�_�����ɒ����_"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^�E"
    .IgnoreCase = False
    .Global = True
  End With
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    slctCells.Value = "�E" & Reg.Replace(slctCells.Value, "")
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ǉ�_�����ɘA��()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.�ǉ�_�����ɘA��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^[0-9]+�D"
    .IgnoreCase = False
    .Global = True
  End With

  line = 1
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    Call Library.showDebugForm("�ݒ�O�Z���l", slctCells.Value, "debug")
    
    slctCells.NumberFormatLocal = "@"
    slctCells.Value = line & "�D" & Reg.Replace(slctCells.Value, "")
    line = line + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �㏑_�����ɘA��()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim i As Long
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.�㏑_�����ɘA��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  line = 1
  
  If Selection.Item(1).Value = "" Then
    line = 1
  Else
    line = Selection.Item(1).Value
  End If
  
  Selection.HorizontalAlignment = xlCenter
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    Call Library.showDebugForm("�ݒ�O�Z���l", slctCells.Value, "debug")
    slctCells.Value = line
    line = line + 1
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function





'==================================================================================================
Function �ϊ�_�S�p�˔��p()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.�ϊ�_�S�p�˔��p"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = Library.convZen2Han(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_���p�ˑS�p()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.�ϊ�_���p�ˑS�p"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = Library.convHan2Zen(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ݒ�_��������()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�ݒ�_��������"
    
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    If slctCells.Font.Strikethrough = True Then
      slctCells.Font.Strikethrough = False
    Else
      slctCells.Font.Strikethrough = True
    End If
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �R�����g_�ǉ�()
  Dim commentVal As String, commentBgColor As Long, CommentFontColor As Long
  Dim CommentFont As String, CommentFontSize As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�R�����g_�ǉ�"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  For Each slctCells In Selection
    
    '��������Ă���ꍇ
    If slctCells.MergeCells Then
      If slctCells.MergeArea.Item(1).Address = slctCells.Address Then
      Else
        GoTo LBl_nextFor
      End If
    End If

    If TypeName(slctCells.Comment) = "Comment" Then
      commentVal = slctCells.Comment.Text
      commentBgColor = slctCells.Comment.Shape.Fill.ForeColor.RGB
      CommentFontSize = slctCells.Comment.Shape.TextFrame.Characters.Font.Size
      CommentFont = slctCells.Comment.Shape.TextFrame.Characters.Font.Name
      CommentFontColor = slctCells.Comment.Shape.TextFrame.Characters.Font.Color
    End If
    
    If FrmVal("commentVal") <> "" Then
      commentVal = FrmVal("commentVal")
    End If
      
    With Frm_InsComment
      .TextBox = commentVal
      If commentVal <> "" Then
        .CommentColor.BackColor = commentBgColor
        .CommentFont = CommentFont
        .CommentFontColor.BackColor = CommentFontColor
        .CommentFontSize = CommentFontSize
      End If
      .Label1.Caption = "�I���Z���F" & slctCells.Address(RowAbsolute:=False, ColumnAbsolute:=False)
      .Show
    End With
    DoEvents

LBl_nextFor:
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �R�����g_�摜�ǉ�()
  Dim insImgPath As String
  
  Const funcName As String = "Ctl_Cells.�R�����g_�摜�ǉ�"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  insImgPath = Library.getFilePath(ActiveWorkbook.path, "", "�R�����g�ɉ摜�}��", "img")
  If insImgPath = "" Then
    GoTo LB_ExitFunction
  End If

  With ActiveCell
    If TypeName(.Comment) = "Comment" Then
      .ClearComments
    End If

    With .AddComment
      .Shape.Fill.UserPicture insImgPath
'      .Shape.Height = 500
'      .Shape.Width = 500
    End With
  End With


  
  '�����I��--------------------------------------
LB_ExitFunction:
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �R�����g_�폜()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�R�����g_�폜"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  If ActiveSheet.ProtectContents = True Then
    Call Library.showNotice(413, , True)
  End If
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, "")

    If TypeName(slctCells.Comment) = "Comment" Then
      slctCells.ClearComments
    End If
    DoEvents
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �\�t_�s�����ւ�()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�\�t_�s�����ւ�"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �㏑_�[��()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�㏑_�[��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    If slctCells.Text = "" Then
      slctCells.Value = 0
      DoEvents
    End If
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �폜_���s()
  Dim slctCells As Range
  Dim resVal As String
  Dim i As Long
  
  Const funcName As String = "Ctl_Cells.�폜_���s"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  arrCells = Selection.Value
  For i = LBound(arrCells, 1) To UBound(arrCells, 1)
    arrCells(i, 1) = Replace(arrCells(i, 1), vbCrLf, "")
    arrCells(i, 1) = Replace(arrCells(i, 1), vbCr, "")
    arrCells(i, 1) = Replace(arrCells(i, 1), vbLf, "")
    
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, UBound(arrCells, 1), "")
  Next
  Selection.Value = arrCells
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �}��_�s()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�}��_�s"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �}��_��()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�}��_��"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �폜_�萔()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.�폜_�萔"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  On Error Resume Next
  If Selection.CountLarge = 1 Then
    Call Library.showNotice(600)
  ElseIf Selection.CountLarge > 1 Then
    Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents
  End If
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function �ϊ�_�啶���ˏ�����()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.�ϊ�_�啶���ˏ�����"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
 
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = LCase(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_�������ˑ啶��()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.�ϊ�_�������ˑ啶��"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0

  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = UCase(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_�ې����ː��l()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.�ϊ�_�ې����ː��l"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    If slctCells.Value <> "" Then
      Call Library.showDebugForm("Value", AscW(slctCells.Value), "debug")
      Select Case AscW(slctCells.Value)
        Case 9450
          slctCells.Value = 0
        
        '1�`20
        Case 9312 To 9332
          slctCells.Value = AscW(slctCells.Value) - 9311
        
        '21�`35
        Case 12881 To 12901
          slctCells.Value = AscW(slctCells.Value) - 12881 + 21
  
        '36�`50
        Case 12977 To 13027
          slctCells.Value = AscW(slctCells.Value) - 12941
          
        Case Else
      End Select
    End If
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_���l�ˊې���()
  Dim line As Long, endLine As Long
  Dim slctCells As Range

  Const funcName As String = "Ctl_Cells.�ϊ�_���l�ˊې���"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    Select Case slctCells.Value
      Case 1 To 20
        slctCells.Value = Chr(Asc("�@") + slctCells.Value - 1)
        
      Case 21 To 35
        slctCells.Value = ChrW(12881 + slctCells.Value - 21)
      
      Case 36 To 50
        slctCells.Value = ChrW(12941 + slctCells.Value)
        
      Case 0
        slctCells.Value = ChrW(9450)
      Case Else
    End Select
  
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ϊ�_URL�G���R�[�h()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�ϊ�_URL�G���R�[�h"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    slctCells.Value = Library.convURLEncode(slctCells.Text)
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_URL�f�R�[�h()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�ϊ�_URL�f�R�[�h"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    slctCells.Value = Library.convURLDecode(slctCells.Text)
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ϊ�_Unicode�G�X�P�[�v()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�ϊ�_Unicode�G�X�P�[�v"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    slctCells.Value = Library.convUnicodeEscape(slctCells.Text)
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ϊ�_Unicode�A���G�X�P�[�v()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�ϊ�_Unicode�A���G�X�P�[�v"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)


    slctCells.Value = Library.convUnicodeunEscape(slctCells.Text)
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
'@link   https://calmdays.net/vbatips/base64encode/
Function �ϊ�_Base64�G���R�[�h()
  Dim slctCells As Range
  Dim oXmlNode  As Object
  
  Const funcName As String = "Ctl_Cells.�ϊ�_Base64�G���R�[�h"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  For Each slctCells In Selection
    Set oXmlNode = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    oXmlNode.dataType = "bin.base64"
    oXmlNode.nodeTypedValue = convStringToBinary(slctCells.Value)
    slctCells.Value = Replace(oXmlNode.Text, "77u/", "")
    Set oXmlNode = Nothing
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'@link   https://calmdays.net/vbatips/base64decode/
Function �ϊ�_Base64�f�R�[�h()
  Dim slctCells As Range
  Dim oXmlNode  As Object
  
  Const funcName As String = "Ctl_Cells.�ϊ�_Base64�f�R�[�h"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Set oXmlNode = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    
    oXmlNode.dataType = "bin.base64"
    oXmlNode.Text = slctCells.Value
    slctCells.Value = convBinaryToString(oXmlNode.nodeTypedValue)
    
    Set oXmlNode = Nothing
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'******************************************************************************
'�@�o�C�i���𕶎���ɕϊ�
'@link   https://calmdays.net/vbatips/base64decode/
'******************************************************************************
Function convBinaryToString(getBinary As Variant) As Variant
  Dim oAdoStm As Object
  Set oAdoStm = CreateObject("ADODB.Stream")

  oAdoStm.Type = 1
  oAdoStm.Open
  oAdoStm.Write getBinary
  oAdoStm.Position = 0
  oAdoStm.Type = 2
  oAdoStm.Charset = "utf-8" 'or "shift_jis"
  convBinaryToString = oAdoStm.ReadText
  
  Set oAdoStm = Nothing
End Function


'******************************************************************************
'�@��������o�C�i���ɕϊ�
'@link   https://calmdays.net/vbatips/base64encode/
'******************************************************************************
Function convStringToBinary(getText As String) As Variant
  Dim oAdoStm As Object
  Set oAdoStm = CreateObject("ADODB.Stream")
  
  oAdoStm.Type = 2
  oAdoStm.Charset = "utf-8" 'or "shift_jis"
  oAdoStm.Open
  oAdoStm.WriteText getText
  oAdoStm.Position = 0
  oAdoStm.Type = 1
  convStringToBinary = oAdoStm.Read
  
  Set oAdoStm = Nothing
End Function
