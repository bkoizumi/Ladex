Attribute VB_Name = "Ctl_UsrForm"
Option Explicit

'**************************************************************************************************
' * 表示位置確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 表示位置(T, L)
  Dim topPosition As Long, leftPosition As Long
  
  topPosition = CLng(T)
  leftPosition = CLng(L)
  
  Call Library.getMachineInfo
  
'  Call Library.showDebugForm("topPosition     ：" & topPosition)
'  Call Library.showDebugForm("leftPosition    ：" & leftPosition)
'  Call Library.showDebugForm("displayX        ：" & MachineInfo("displayX"))
'  Call Library.showDebugForm("displayY        ：" & MachineInfo("displayY"))
'  Call Library.showDebugForm("displayVirtualX ：" & MachineInfo("displayVirtualX"))
'  Call Library.showDebugForm("displayVirtualY ：" & MachineInfo("displayVirtualY"))
'
'  Call Library.showDebugForm("appWidth ：" & MachineInfo("appWidth"))
'  Call Library.showDebugForm("appHeight ：" & MachineInfo("appHeight"))

  
  If topPosition > MachineInfo("appHeight") Then
    T = CInt(MachineInfo("appHeight") / 4)
  ElseIf topPosition = 0 Then
    T = CInt(MachineInfo("appHeight") / 4)
  Else
    T = topPosition
  End If
  
  If leftPosition > MachineInfo("appWidth") Then
    L = CInt(MachineInfo("appWidth") / 4)
  ElseIf leftPosition = 0 Then
    L = CInt(MachineInfo("appWidth") / 4)
  Else
    L = leftPosition
  End If
  
'  Call Library.showDebugForm("t               ：" & t)
'  Call Library.showDebugForm("l               ：" & l)


End Function



'**************************************************************************************************
' * イベント処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 日付(inputVal As Variant)

'  Call Library.showDebugForm("inputVal：" & inputVal)
  
  If IsDate(inputVal) Then
    inputVal = Format(inputVal, "yyyy/mm/dd")
  ElseIf inputVal = "" Then
    inputVal = ""
  Else
    inputVal = False
  End If
  
  日付 = inputVal
  
End Function

