Attribute VB_Name = "Ctl_�J�X�^��"
Option Explicit

'==================================================================================================
Function �J�X�^���֐�01()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_�J�X�^��.�J�X�^���֐�01"

  '�����J�n--------------------------------------
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
  
  Dim SelectionSheetName As String
  
  SelectionSheetName = ActiveSheet.Previous.Name
  ActiveSheet.Move After:=Sheets(ActiveWorkbook.Worksheets.count)
  ActiveWorkbook.Worksheets(SelectionSheetName).Select
  ActiveWorkbook.ActiveSheet.Next.Activate


  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function �J�X�^���֐�02()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_�J�X�^��.�J�X�^���֐�02"

  '�����J�n--------------------------------------
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
  
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'  MsgBox (funcName)


  Selection.ShapeRange.Item(1).Width = 236


  '�����I��--------------------------------------
  If runFlg = False Then
    'Application.GoTo Reference:=Range("A1"), Scroll:=True
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
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function �J�X�^���֐�03()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_�J�X�^��.�J�X�^���֐�03"

  '�����J�n--------------------------------------
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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)





  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function �J�X�^���֐�04()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_�J�X�^��.�J�X�^���֐�04"

  '�����J�n--------------------------------------
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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)





  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function �J�X�^���֐�05()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_�J�X�^��.�J�X�^���֐�05"

  '�����J�n--------------------------------------
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
  endLine = Cells(Rows.count, 1).End(xlUp).Row


  MsgBox (funcName)




  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function






