Attribute VB_Name = "Ctl_HighLight"
'Option Explicit
'Option Private Module

#If VBA7 And Win64 Then

#Else
    
  Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

  Public Declare Function GetCursorPos Lib "user32" (IpPoint As POINTAPI) As Long

#End If

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_LAYERED = &H80000
Public Const WS_CAPTION = &HC00000
Public Const WS_EX_DLGMODALFRAME = &H1&

Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Const DPI As Long = 96    'dots per inch
Public Const PPI As Long = 72    'pixel per inch
  

'==================================================================================================
Function showStart(ByVal Target As Range)
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  Set Rng = Range("A" & Target.Row)
  Set Rng = Target
  Call init.setting
  
  Call Library.getCellPosition(Rng, ActvCellTop, ActvCellLeft)
  
  If Frm_HighLight.Visible Then
    Unload Frm_HighLight
  End If
  
  With Frm_HighLight
    .StartUpPosition = 0
    .Top = ActvCellTop
    .Left = ActvCellLeft
    .Width = Application.Width
    .Height = Rng.Height
    
    .BackColor = Library.getRegistry(RegistrySubKey, "HighLightColor")
    
    .Show
  End With
  Set Rng = Nothing


End Function


'==================================================================================================
Function showEnd()
  
  Unload Frm_HighLight
  
End Function

