Attribute VB_Name = "Ctl_UsrForm"
Option Explicit
 
Private Const GWL_STYLE = -16
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
 
#If Win64 Then
  Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
  Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
  Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
  
  Dim hwnd As LongPtr
  Dim rc As LongPtr

#Else
  Declare Function GetActiveWindow Lib "user32" () As Long
  Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA"(ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA"(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
  
  Dim hwnd As Long
  Dim rc As Long
    
#End If
'**************************************************************************************************
' * �T�C�Y���ω�����
' *
' * @Link   https://liclog.net/setwindowlong-function-vba-api/
'**************************************************************************************************
'==================================================================================================
Function ResizeForm()
  Dim style As Long
 
  hwnd = GetActiveWindow()
  
  '�擾�����E�C���h�E�̃X�^�C�����擾
  style = GetWindowLong(hwnd, GWL_STYLE)
  
  '�擾�����E�C���h�E�̃X�^�C���ɃT�C�Y�ρ{�ő剻�{�^���{�ŏ����{�^���ǉ�
  style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
  rc = SetWindowLong(hwnd, GWL_STYLE, style)
  
  '�E�C���h�E�̃X�^�C�����ĕ`��
  rc = DrawMenuBar(hwnd)
  
End Function


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

