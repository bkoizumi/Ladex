Attribute VB_Name = "Ctl_カスタム"
Option Explicit

'==================================================================================================
Function カスタム01()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_カスタム.カスタム01"

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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)




  '処理終了--------------------------------------
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

  'エラー発生時------------------------------------
catchError:
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function カスタム02()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_カスタム.カスタム02"

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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)





  '処理終了--------------------------------------
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

  'エラー発生時------------------------------------
catchError:
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function カスタム03()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_カスタム.カスタム03"

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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)





  '処理終了--------------------------------------
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

  'エラー発生時------------------------------------
catchError:
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function カスタム04()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_カスタム.カスタム04"

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
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  MsgBox (funcName)





  '処理終了--------------------------------------
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

  'エラー発生時------------------------------------
catchError:
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function カスタム05()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_カスタム.カスタム05"

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
  endLine = Cells(Rows.count, 1).End(xlUp).Row


  MsgBox (funcName)




  '処理終了--------------------------------------
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

  'エラー発生時------------------------------------
catchError:
    Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function






