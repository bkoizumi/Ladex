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
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
'// 32bit版
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If


'**************************************************************************************************
' * フォームサイズ変更
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function FormResize()
  Dim hWnd  As Long, style As Long
  
  'ウインドウハンドル取得
  hWnd = GetActiveWindow()
  
  'ウインドウのスタイルを取得
  style = GetWindowLong(hWnd, GWL_STYLE)
  
  'ウインドウのスタイルにウインドウサイズ可変＋最小ボタン＋最大ボタンを追加
  style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
  
  'ウインドウのスタイルを再設定
  Call SetWindowLong(hWnd, GWL_STYLE, style)
End Function


'**************************************************************************************************
' * ユーザー定義関数
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeUsrFunction()
  Const funcName As String = "Ctl_UsrFunction.InitializeUsrFunction"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
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
    ArgumentDescriptions:=Array("チェックする日付を指定", "第N週を数値で指定", "曜日を数値で指定" & vbNewLine & _
                                "1：月　2：火　3：水" & vbNewLine & _
                                "4：木　5：金　6：土　7：日")

  Application.MacroOptions _
    Macro:="Textjoin", _
    Description:="文字列連結", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("区切り文字", "空欄時処理[True：処理する/False：処理しない]", "文字列1,文字列2, ...")

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function Textjoin(Delim As String, ParamArray par())
Attribute Textjoin.VB_Description = "文字列連結"
Attribute Textjoin.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim i As Integer
  Dim tR As Range

'  Application.Volatile

  Textjoin = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Then
          Textjoin = Textjoin & Delim & tR.Value2
        End If
      Next
    Else
      If (par(i) <> "" And par(i) <> "<>") Then
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


'==================================================================================================
'Function mkQRcode(ByVal codeVal As String, Optional ByVal QRSize As Long = 140) As String
'  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  Dim slctCells As Range
'
'  Dim chartAPIURL As String
'  Dim QRCodeImgName As String
'
'  Const funcName As String = "Ctl_Shap.QRコード生成"
'  Const chartAPI = "https://chart.googleapis.com/chart?cht=qr&chld=l|1&"
'
'
'  '処理開始--------------------------------------
''  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
'  '----------------------------------------------
'
'  Set slctCells = Application.Caller
'  QRCodeImgName = "QRCode_" & Application.Caller.Address(False, False)
'
'  Call Library.showDebugForm("slctCells", slctCells.Address, "debug")
'  Call Library.showDebugForm("codeVal", codeVal, "debug")
'  Call Library.showDebugForm("QRCodeImgName", QRCodeImgName, "debug")
'
'
'  If Library.chkShapeName(QRCodeImgName) Then
'    ActiveSheet.Shapes.Range(Array(QRCodeImgName)).Select
'    Selection.delete
'  End If
'
'
'  chartAPIURL = chartAPI & "chs=" & QRSize & "x" & QRSize
'  chartAPIURL = chartAPIURL & "&chl=" & Library.convURLEncode(codeVal)
'  Call Library.showDebugForm("chartAPIURL", chartAPIURL, "debug")
'
'  With ActiveSheet.Pictures.Insert(chartAPIURL)
'    .ShapeRange.Top = slctCells.Top + (slctCells.Height - .ShapeRange.Height) / 2
'    .ShapeRange.Left = slctCells.Left + (slctCells.Width - .ShapeRange.Width) / 2
'
'    .Placement = xlMove
'
'    'QRコードの名前設定
'    .ShapeRange.Name = QRCodeImgName
'    .Name = QRCodeImgName
'  End With
'
'  mkQRcode = ""
''  ActiveSheet.Select
''  slctCells.Select
'  Set slctCells = Nothing
'
'  '処理終了--------------------------------------
'Lbl_endFunction:
'  Call Library.endScript
'  Call Library.showDebugForm(funcName, , "end")
'  Call init.unsetting
'  '----------------------------------------------
'  Exit Function
'
''エラー発生時------------------------------------
'catchError:
'  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
'End Function


