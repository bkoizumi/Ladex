Attribute VB_Name = "Ctl_System"
#If VBA7 And Win64 Then
Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

#Else
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
#End If

'**************************************************************************************************
' * �V�X�e���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�R���g���[���p�l���̃z�C�[���ʎ擾
Function getScroll()
  Dim scrollVal As Long
  Const GetWheelScrollLines = 104
  Const funcName As String = "Ctl_System.getScroll"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SystemParametersInfo GetWheelScrollLines, 0, scrollVal, 0
  Range("scrollRowCnt") = scrollVal
  
  
  
  'Call Library.showDebugForm("getScroll�F" & Range("scrollVal"))
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'�R���g���[���p�l���̃z�C�[���ʂP�s
Function setScroll()

  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
    

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------



  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
On Error GoTo catchError
  
  Call init.setting
  SystemParametersInfo SetWheelScrollLines, 1, 0, SENDCHANGE
  'Call Library.showDebugForm("setScroll�F" & 1)
  
  Exit Function
'�G���[������------------------------------------
catchError:
  'Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
'�R���g���[���p�l���̃z�C�[���ʂ�߂�
Function resetScroll()

  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
  

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------



  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
On Error GoTo catchError
  
  Call init.setting
  SystemParametersInfo SetWheelScrollLines, Range("scrollVal"), 0, SENDCHANGE
  'Call Library.showDebugForm("setScroll�F" & Range("scrollVal"))
  
  Exit Function
'�G���[������------------------------------------
catchError:
  'Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function
