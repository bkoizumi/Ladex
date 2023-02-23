Attribute VB_Name = "Ctl_System"
#If VBA7 And Win64 Then
Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

#Else
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
#End If

'**************************************************************************************************
' * システム設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'コントロールパネルのホイール量取得
Function getScroll()
  Dim scrollVal As Long
  Const GetWheelScrollLines = 104
  Const funcName As String = "Ctl_System.getScroll"

  '処理開始--------------------------------------
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
  
  
  
  
  '処理終了--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

  getScroll = scrollVal
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'コントロールパネルのホイール量１行
Function setScroll(Optional scrollVal As Integer = 1)
  Const funcName As String = "Ctl_System.setScroll"
  
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
    

  '処理開始--------------------------------------
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
    
    

  '処理終了--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
'コントロールパネルのホイール量を戻す
Function resetScroll()
  Const funcName As String = "Ctl_System.resetScroll"
  
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
  

  '処理開始--------------------------------------
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

  '処理終了--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  'Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function
