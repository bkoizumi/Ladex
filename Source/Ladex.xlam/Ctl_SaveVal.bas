Attribute VB_Name = "Ctl_SaveVal"
Option Explicit

'**************************************************************************************************
' * VBA実行前の値を保持
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setVal(pType As String, pText As String)
  Dim line As Long, endLine As Long
  Dim chkFlg As Boolean
  Const funcName As String = "Ctl_SaveVal.setVal"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  chkFlg = False
  LadexBook.Activate
  
  endLine = Cells(Rows.count, 4).End(xlUp).Row
  For line = 3 To LadexSh_Config.Cells(Rows.count, 4).End(xlUp).Row
    If LadexSh_Config.Range(dicVal("Cells_pType") & line) = pType Then
      LadexSh_Config.Range(dicVal("Cells_pType") & line) = pType
      LadexSh_Config.Range(dicVal("Cells_pText") & line) = pText
      
      chkFlg = True
      Exit For
    End If
  Next
   
  If chkFlg = False Then
    LadexSh_Config.Range(dicVal("Cells_pType") & line) = pType
    LadexSh_Config.Range(dicVal("Cells_pText") & line) = pText
  End If
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
End Function


'==================================================================================================
Function getVal(pType As String) As String
  Dim resetObjVal          As Object
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SaveVal.getVal"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Set resetObjVal = Nothing
  Set resetObjVal = CreateObject("Scripting.Dictionary")
  
  For line = 3 To LadexSh_Config.Cells(Rows.count, 4).End(xlUp).Row
    If LadexSh_Config.Range(dicVal("Cells_pType") & line) <> "" Then
      resetObjVal.add LadexSh_Config.Range(dicVal("Cells_pType") & line).Text, LadexSh_Config.Range(dicVal("Cells_pText") & line).Text
    End If
  Next
  
  If resetObjVal("reSet" & pType) = "" Then
    getVal = resetObjVal(pType)
  Else
    getVal = resetObjVal("reSet" & pType)
  End If
  Set resetObjVal = Nothing
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
End Function

'==================================================================================================
Function delVal(pType As String)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SaveVal.delVal"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For line = 3 To LadexSh_Config.Cells(Rows.count, 4).End(xlUp).Row
    If LadexSh_Config.Range(dicVal("Cells_pType") & line) Like "*" & pType Then
      LadexSh_Config.Range(dicVal("Cells_pType") & line) = ""
      LadexSh_Config.Range(dicVal("Cells_pText") & line) = ""
    End If
  Next
 
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
End Function

