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
'  If runFlg = False Then
'    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
'    Call Library.startScript
'  Else
'    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
'  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SystemParametersInfo GetWheelScrollLines, 0, scrollVal, 0
  
  Call Library.showDebugForm("scrollVal", scrollVal, "debug")
  LadexSetVal("scrollVal") = scrollVal
  
  
  
  
  '�����I��--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

  getScroll = scrollVal
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'�R���g���[���p�l���̃z�C�[���ʂP�s
Function setScroll(Optional scrollVal As Integer = 1)
  Const funcName As String = "Ctl_System.setScroll"
  
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
    

  '�����J�n--------------------------------------
'  If runFlg = False Then
'    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
'    Call Library.startScript
'  Else
'    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
'  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
    
  SystemParametersInfo SetWheelScrollLines, 1, 0, SENDCHANGE
  Call Library.showDebugForm("scrollVal", 1, "debug")
    
    

  '�����I��--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
'�R���g���[���p�l���̃z�C�[���ʂ�߂�
Function resetScroll()
  Const funcName As String = "Ctl_System.resetScroll"
  
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
  

  '�����J�n--------------------------------------
'  If runFlg = False Then
'    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
'    Call Library.startScript
'  Else
'    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
'  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  LadexSetVal("scrollVal") = Library.getRegistry("Main", "scrollVal")
  
  SystemParametersInfo SetWheelScrollLines, LadexSetVal("scrollVal"), 0, SENDCHANGE

  Call Library.showDebugForm("scrollVal", LadexSetVal("scrollVal"), "debug")

  '�����I��--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  'Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function
