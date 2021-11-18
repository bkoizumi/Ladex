Attribute VB_Name = "Ctl_UsrFunction"
Option Explicit

'// Win32API用定数
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
'// Win32API参照宣言
'// 64bit版
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
'// 32bit版
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If


'**************************************************************************************************
' * フォームサイズ変更
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function FormResize()
    Dim hwnd                    As Long         '// ウインドウハンドル
    Dim style                   As Long         '// ウインドウスタイル
 
    '// ウインドウハンドル取得
    hwnd = GetActiveWindow()
    
    '// ウインドウのスタイルを取得
    style = GetWindowLong(hwnd, GWL_STYLE)
    
    '// ウインドウのスタイルにウインドウサイズ可変＋最小ボタン＋最大ボタンを追加
    style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
 
    '// ウインドウのスタイルを再設定
    Call SetWindowLong(hwnd, GWL_STYLE, style)
End Function


'**************************************************************************************************
' * ユーザー定義関数
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeUsrFunction()

  Application.MacroOptions _
    Macro:="chkWorkDay", _
    Description:="第N営業日かチェックし、True/Falseを返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする日付を指定", "第N営業日を数値で指定")

  Application.MacroOptions _
    Macro:="getWorkDay", _
    Description:="第N営業日をシリアル値で返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする年を数値で指定", "チェックする月を数値で指定", "第N営業日を数値で指定")

  Application.MacroOptions _
    Macro:="chkWeekNum", _
    Description:="第N週X曜日の日付かチェックし、True/Falseを返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする日付を指定", "第N週を数値で指定", "X曜日を数値で指定")

  Application.MacroOptions _
    Macro:="Textjoin", _
    Description:="文字列連結", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("区切り文字", "空欄時処理[True：処理する/False：処理しない]", "文字列1,文字列2, ...")


End Function



'==================================================================================================
Function Textjoin(Delim, Ignore As Boolean, ParamArray par())
Attribute Textjoin.VB_Description = "文字列連結"
Attribute Textjoin.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim i As Integer
  Dim tR As Range

'  Application.Volatile

  Textjoin = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Or Ignore = False Then
          Textjoin = Textjoin & Delim & tR.Value2
        End If
      Next
    Else
      If (par(i) <> "" And par(i) <> "<>") Or Ignore = False Then
        Textjoin = Textjoin & Delim & par(i)
      End If
    End If
  Next

  Textjoin = Mid(Textjoin, Len(Delim) + 1)

End Function




'==================================================================================================
Function chkWorkDay(ByVal checkDate As Date, ByVal bsnDay As Long) As Boolean
Attribute chkWorkDay.VB_Description = "第N営業日かチェックし、True/Falseを返す"
Attribute chkWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
  
'  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(Year(checkDate), Month(checkDate), 1)
  getDay = Application.WorksheetFunction.WorkDay(firstDay, bsnDay, arryHollyday)
  
  If checkDate = getDay Then
    chkWorkDay = True
  Else
    chkWorkDay = False
  End If

End Function

'==================================================================================================
Function chkWeekNum(ByVal checkDate As Date, ByVal checkWeekday As Long, ByVal weekNum As Long) As Boolean
Attribute chkWeekNum.VB_Description = "第N週X曜日の日付かチェックし、True/Falseを返す"
Attribute chkWeekNum.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Long, diff As Long
  
'  Application.Volatile
  
  firstDay = Weekday(DateSerial(Year(checkDate), Month(checkDate), 1))
  diff = (checkWeekday + 7 - firstDay) Mod 7
  getDay = DateSerial(Year(checkDate), Month(checkDate), 1 + diff + 7 * (weekNum - 1))
  
  If checkDate = getDay Then
    chkWeekNum = True
  Else
    chkWeekNum = False
  End If
  
End Function

'==================================================================================================
Function getWorkDay(ByVal cYear As Long, ByVal cMonth As Long, ByVal bsnDay As Long) As Date
Attribute getWorkDay.VB_Description = "第N営業日をシリアル値で返す"
Attribute getWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
'  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(cYear, cMonth, 1)
  getWorkDay = Application.WorksheetFunction.WorkDay(firstDay - 1, bsnDay, arryHollyday)
  
End Function

