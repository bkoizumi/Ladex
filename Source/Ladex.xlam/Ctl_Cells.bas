Attribute VB_Name = "Ctl_Cells"
Option Explicit

'**************************************************************************************************
' * �Z������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �Z�������Œ�ݒ�(setRowHeight As Integer)
  Dim line As Long, startLine As Long, endLine As Long
  Dim SelectionCell As Range
  
  Const funcName As String = "Ctl_Cells.�Z�������Œ�ݒ�"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")
  Call Library.showDebugForm("setRowHeight", setRowHeight, "debug")

  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  startLine = Selection(1).Row
  endLine = Selection(Selection.count)
  If endLine = 0 Then
    endLine = startLine
  End If
  Rows(startLine & ":" & endLine).rowHeight = setRowHeight
  
  
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function





'==================================================================================================
Function �O��̃X�y�[�X�폜()
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.�O��̃X�y�[�X�폜"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �S�󔒍폜()
  Dim slctCells As Range
  Dim resVal As String
  
  Const funcName As String = "Ctl_Cells.Trim01"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �����_�t�^()
  Dim line As Long, endLine As Long
  Dim Reg As Object
  Dim slctCells
  
  Const funcName As String = "Ctl_Cells.�����_�t�^"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �A�Ԑݒ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim i As Long
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.�A�Ԑݒ�"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �A�Ԓǉ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.�A�Ԓǉ�"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �p�����S�˔��p�ϊ�()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.�p�����S�˔��p�ϊ�"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �p�������ˑS�p�ϊ�()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.�p�������ˑS�p�ϊ�"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���������ݒ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.���������ݒ�"
    
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �R�����g�}��()
  Dim commentVal As String, commentBgColor As Long, CommentFontColor As Long
  Dim CommentFont As String, CommentFontSize As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�R�����g�}��"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �R�����g�폜()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�R�����g�폜"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �s������ւ��ē\�t��()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�s������ւ��ē\�t��"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �[������()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�[������"
  
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���s�폜()
  Dim slctCells As Range
  Dim resVal As String
  Dim i As Long
  
  Const funcName As String = "Ctl_Cells.���s�폜"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �s�}��()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.�s�}��"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ��}��()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.��}��"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �萔�폜()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.�萔�폜"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z��������()
  Dim colLine As Long, endColLine As Long
  Dim slctStartColLine As Long, slctEndColLine As Long, useEndColLine As Long
  Dim colName As String
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.�Z��������"
  Const maxColumnWidth As Integer = 60
  
  
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
  
  useEndColLine = Selection.SpecialCells(xlLastCell).Column
  
  '�I��͈͂�����ꍇ----------------------------
  If Selection.CountLarge > 1 Then
    Call Library.showDebugForm("�I��͈͂���", , "debug")
    
    Selection.EntireColumn.AutoFit
  
    slctStartColLine = Selection.Column
    slctEndColLine = Selection.Column + Selection.Columns.count
  
  '�I��͈͂��Ȃ��ꍇ----------------------------
  Else
    Call Library.showDebugForm("�I��͈͂Ȃ�", , "debug")
    Cells.EntireColumn.AutoFit
  
    slctStartColLine = 1
    slctEndColLine = useEndColLine
  
  End If
  
  '�ő啝�m�F------------------------------------
  For colLine = slctStartColLine To slctEndColLine
    colName = Library.getColumnName(colLine)
    Call Library.showDebugForm("�ő啝�m�F", colName, "debug")
    
    
    If Cells(1, colLine).ColumnWidth > maxColumnWidth Then
      Columns(colName & ":" & colName).ColumnWidth = maxColumnWidth
    Else
      Columns(colName & ":" & colName).ColumnWidth = WorksheetFunction.RoundUp(Columns(colName & ":" & colName).ColumnWidth, 0) + 1
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �Z����������()
  Dim line As Long, startLine As Long, endLine As Long
  Dim SelectionCell As Range
  
  Const funcName As String = "Ctl_Cells.�Z����������"
  
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
  
  Cells.EntireRow.AutoFit
  
  If Selection.Rows.count <= 1 Then
    startLine = 1
    endLine = Range("A1").SpecialCells(xlLastCell).Row
  Else
    startLine = SelectionCell.Row
    endLine = Range("A1").SpecialCells(xlLastCell).Row
  End If
  Selection.EntireRow.AutoFit
  
  
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "�����ݒ�")
    
    If Rows(line & ":" & line).Hidden = False Then
      If Rows(line & ":" & line).Height < Int(dicVal("rowHeight")) Then
        Rows(line & ":" & line).rowHeight = dicVal("rowHeight")
      
      ElseIf Rows(line & ":" & line).Height > 200 Then
        Rows(line & ":" & line).rowHeight = 200
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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ��ˏ��ϕϊ�()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.��ˏ��ϕϊ�"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���ˑ�ϕϊ�()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.���ˑ�ϕϊ�"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ې����ː��l()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.�ې����ː��l"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���l�ˊې���()
  Dim line As Long, endLine As Long
  Dim slctCells As Range

  Const funcName As String = "Ctl_Cells.���l�ˊې���"

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
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function URL�G���R�[�h()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.URL�G���R�[�h"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function URL�f�R�[�h()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.URL�f�R�[�h"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function Unicode�G�X�P�[�v()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.Unicode�G�X�P�[�v"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function Unicode�A���G�X�P�[�v()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.Unicode�A���G�X�P�[�v"

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

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
