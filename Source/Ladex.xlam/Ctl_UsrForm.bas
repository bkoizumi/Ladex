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
' * サイズを可変化する
' *
' * @Link   https://liclog.net/setwindowlong-function-vba-api/
'**************************************************************************************************
'==================================================================================================
Function ResizeForm()
  Dim style As Long
 
  hwnd = GetActiveWindow()
  
  '取得したウインドウのスタイルを取得
  style = GetWindowLong(hwnd, GWL_STYLE)
  
  '取得したウインドウのスタイルにサイズ可変＋最大化ボタン＋最小化ボタン追加
  style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
  rc = SetWindowLong(hwnd, GWL_STYLE, style)
  
  'ウインドウのスタイルを再描画
  rc = DrawMenuBar(hwnd)
  
End Function


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

