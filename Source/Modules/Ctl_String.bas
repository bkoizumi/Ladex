Attribute VB_Name = "Ctl_String"
Option Explicit

'**************************************************************************************************
' * �����񑀍�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Trim01()
  ActiveCell = Trim(ActiveCell.Text)
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Function

'==================================================================================================
Function �����_�t�^()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.�����_�t�^"

  Call Library.startScript
  Call init.setting
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  For Each slctCells In Selection
    slctCells.Value = "�E" & slctCells.Value
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'==================================================================================================
Function �A�ԕt�^()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.�A�ԕt�^"

  Call Library.startScript
  Call init.setting
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
    slctCells.Value = line & "�D" & slctCells.Value
    line = line + 1
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function

