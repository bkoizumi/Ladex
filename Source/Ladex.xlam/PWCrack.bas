Attribute VB_Name = "PWCrack"
'**************************************************************************************************
' * VBA-マクロのパスワード解除
' *
' * @Link   https://nkmrdai.com/vba-password-unrocked/
'**************************************************************************************************


'Win32API参照宣言
'64bit版
#If VBA7 And Win64 Then
  Public Const PAGE_EXECUTE_READWRITE = &H40
  Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
  Public Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
  Public Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
  Public Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
  Public Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
  
  Dim HookBytes(0 To 5) As Byte
  Dim OriginBytes(0 To 5) As Byte
  Dim projectFunction As LongPtr
  Dim Flag As Boolean
 
'32bit版
#Else
  Public Const PAGE_EXECUTE_READWRITE = &H40
   
  Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
  Public Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
  Public Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
  Public Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
  Public Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
  
  Dim HookBytes(0 To 5) As Byte
  Dim OriginBytes(0 To 5) As Byte
  Dim projectFunction As Long
  Dim Flag As Boolean
#End If
 
'==================================================================================================
Sub VBAProjectパスワード解除()
  If HookFlag Then
    MsgBox "VBA Project を解除しました。", vbInformation, "成功しました。"
  Else
    MsgBox "VBA Project 解除に失敗しました。"
  End If
End Sub

'==================================================================================================
Public Function GetPtr(ByVal Value As LongPtr) As LongPtr
  GetPtr = Value
End Function
 
'==================================================================================================
Public Sub RecoverBytes()
  If Flag Then MoveMemory ByVal projectFunction, ByVal VarPtr(OriginBytes(0)), 6
End Sub

'==================================================================================================
Public Function MyDialogBoxParamater(ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
  If pTemplateName = 4070 Then
    MyDialogBoxParamater = 1
  Else
    RecoverBytes
    MyDialogBoxParamater = MyDialogBoxParamater(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam)
    HookFlag
  End If
End Function
 
'==================================================================================================
Public Function HookFlag() As Boolean
  Dim TmpBytes(0 To 5) As Byte
  
#If VBA7 And Win64 Then
  Dim p As LongPtr
  Dim OriginProtect As LongPtr
#Else
  Dim p As Long
  Dim OriginProtect As Long
#End If

  HookFlag = False
  projectFunction = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
  If VirtualProtect(ByVal projectFunction, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
    MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal projectFunction, 6
    If TmpBytes(0) <> &H68 Then
      MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal projectFunction, 6
      p = GetPtr(AddressOf MyDialogBoxParamater)
      HookBytes(0) = &H68
      MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
      HookBytes(5) = &HC3
      MoveMemory ByVal projectFunction, ByVal VarPtr(HookBytes(0)), 6
      Flag = True
      HookFlag = True
    End If
  End If
End Function
 
 

