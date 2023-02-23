Attribute VB_Name = "Ctl_Assign"
Option Explicit

'==================================================================================================
Function �S���҃��X�g�\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim i As Integer
  Const funcName As String = "Ctl_Assign.�S���҃��X�g�\��"

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
  

  '�S���ғǂݍ���--------------------------------
  Erase lstAssign
  endLine = sh_Option.Cells(Rows.count, 11).End(xlUp).Row
  i = 0
  On Error Resume Next
  For line = 4 To endLine
    If sh_Option.Range("L" & line) <> "" Then
      ReDim Preserve lstAssign(i)
      
      lstAssign(i) = sh_Option.Range("K" & line).Text
      Call Library.showDebugForm("�S���ҁF", sh_Option.Range("K" & line).Text, "debug")
      
      i = i + 1
    End If
  Next
  
  




  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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
