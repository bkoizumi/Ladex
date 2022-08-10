Attribute VB_Name = "Ctl_UsrForm"
Option Explicit

'**************************************************************************************************
' * �\���ʒu�m�F
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �\���ʒu(T, L)
  Dim topPosition As Long, leftPosition As Long
  
  topPosition = CLng(T)
  leftPosition = CLng(L)
  
  Call Library.getMachineInfo
  
'  Call Library.showDebugForm("topPosition     �F" & topPosition)
'  Call Library.showDebugForm("leftPosition    �F" & leftPosition)
'  Call Library.showDebugForm("displayX        �F" & MachineInfo("displayX"))
'  Call Library.showDebugForm("displayY        �F" & MachineInfo("displayY"))
'  Call Library.showDebugForm("displayVirtualX �F" & MachineInfo("displayVirtualX"))
'  Call Library.showDebugForm("displayVirtualY �F" & MachineInfo("displayVirtualY"))
'
'  Call Library.showDebugForm("appWidth �F" & MachineInfo("appWidth"))
'  Call Library.showDebugForm("appHeight �F" & MachineInfo("appHeight"))

  
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
  
'  Call Library.showDebugForm("t               �F" & t)
'  Call Library.showDebugForm("l               �F" & l)


End Function



'**************************************************************************************************
' * �C�x���g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function ���t(inputVal As Variant)

'  Call Library.showDebugForm("inputVal�F" & inputVal)
  
  If IsDate(inputVal) Then
    inputVal = Format(inputVal, "yyyy/mm/dd")
  ElseIf inputVal = "" Then
    inputVal = ""
  Else
    inputVal = False
  End If
  
  ���t = inputVal
  
End Function

