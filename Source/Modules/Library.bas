Attribute VB_Name = "Library"
Option Explicit

'**************************************************************************************************
' * 参照設定、定数宣言
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' 利用する参照設定まとめ
' Microsoft Office 14.0 Object Library
' Microsoft DAO 3.6 Objects Library
' Microsoft Scripting Runtime (WSH, FileSystemObject)
' Microsoft ActiveX Data Objects 2.8 Library
' UIAutomationClient

' Windows APIの利用--------------------------------------------------------------------------------
' ディスプレイの解像度取得用
' Sleep関数の利用
' クリップボード関数の利用
#If VBA7 And Win64 Then
  'ディスプレイの解像度取得用
  Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'Sleep関数の利用
  Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As LongPtr)

  'クリップボード関連
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

#Else
  'ディスプレイの解像度取得用
  Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


  'Sleep関数の利用
  Private Declare Function Sleep Lib "kernel32" (ByVal ms As Long)

  'クリップボード関連
  Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function EmptyClipboard Lib "user32" () As Long


  'Shell関数で起動したプログラムの終了を待つ
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Private Const PROCESS_QUERY_INFORMATION = &H400&
  Private Const STILL_ACTIVE = &H103&

#End If



'ワークブック用変数------------------------------
'ワークシート用変数------------------------------
'グローバル変数----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

'アクティブセルの取得
Dim SelectionCell As String
Dim SelectionSheet As String

' PC、Office等の情報取得用連想配列
Public MachineInfo As Object

' Selenium用設定
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

'Public Const JapaneseCharacters = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ"
'Public Const JapaneseCharactersCommonUse = "雨学空金青林画岩京国姉知長直店東歩妹明門夜委育泳岸苦具幸始使事実者昔取受所注定波板表服物放味命油和英果芽官季泣協径固刷参治周松卒底的典毒念府法牧例易往価河居券効妻枝舎述承招性制版肥非武沿延拡供呼刻若宗垂担宙忠届乳拝並宝枚依押奇祈拠況屈肩刺沼征姓拓抵到突杯泊拍迫彼怖抱肪茂炎欧殴"
'Public Const MachineDependentCharacters = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㍉纊褜鍈銈蓜俉炻昱棈鋹曻彅丨仡仼伀伃伹佖侒侊侚侔俍偀倢俿倞偆偰偂傔"

Public ThisBook As Workbook


'**************************************************************************************************
' * アドオンを閉じる
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function addinClose()
  Workbooks(ThisWorkbook.Name).Close
End Function


'**************************************************************************************************
' * 画面描写制御開始
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()
  
  Const funcName As String = "Library.startScript"
  
  On Error Resume Next
  'アクティブセルの取得
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
    SelectionSheet = ActiveWorkbook.ActiveSheet.Name
  End If

  'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = False

  'マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = False

  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationManual

  'マクロ動作中に一切のキーやマウス操作を制限する
  'Application.Interactive = False

  'マクロ動作中はマウスカーソルを「砂時計」にする
  'Application.Cursor = xlWait

  '確認メッセージを出さない
  Application.DisplayAlerts = False

  'Call Library.showDebugForm("funcName", funcName)
End Function


'**************************************************************************************************
' * 画面描写制御終了
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional reCalflg As Boolean = False, Optional flg As Boolean = False)
  Const funcName As String = "Library.endScript"
  
  On Error Resume Next
  
  '強制的に再計算させる
  If reCalflg = True Then
    Application.CalculateFull
  End If

 'アクティブセルの選択
  If SelectionCell <> "" And flg = True Then
    ActiveWorkbook.Worksheets(SelectionSheet).Select
    ActiveWorkbook.Range(SelectionCell).Select
  End If
'  Call unsetClipboard

  'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = True

  'マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = True

  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationAutomatic

  'マクロ動作中に一切のキーやマウス操作を制限する
  'Application.Interactive = True

  'マクロ動作終了後はマウスカーソルを「デフォルト」にもどす
  Application.Cursor = xlDefault

  'マクロ動作終了後はステータスバーを「デフォルト」にもどす
  Application.StatusBar = False

  '確認メッセージを出さない
  Application.DisplayAlerts = True
  
  'Call Library.showDebugForm("funcName", funcName)
End Function


'**************************************************************************************************
' * シートの存在確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkSheetExists(sheetName) As Boolean
  Dim tempSheet As Object
  Dim result As Boolean

  result = False
  For Each tempSheet In Sheets
    If LCase(sheetName) = LCase(tempSheet.Name) Then
      result = True
      Exit For
    End If
  Next
  chkSheetExists = result
End Function


'**************************************************************************************************
' * 処理完了まで待機
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShellEnd(ProcessID As Long)
  Dim hProcess As Long
  Dim EndCode As Long
  Dim EndRet   As Long

  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, ProcessID)
  Do
    EndRet = GetExitCodeProcess(hProcess, EndCode)
    DoEvents
  Loop While (EndCode = STILL_ACTIVE)
  EndRet = CloseHandle(hProcess)
End Function


'**************************************************************************************************
' * オートシェイプの存在確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShapeName(ShapeName As String) As Boolean
  Dim objShp As Shape
  Dim result As Boolean

  result = False
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name = ShapeName Then
      result = True
      Exit For
    End If
  Next
  chkShapeName = result
End Function


'**************************************************************************************************
' * 除外シート判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(chkSheetName As String) As Boolean
  Dim result As Boolean
  Dim sheetName As Variant

  For Each sheetName In Range("ExcludeSheet")
    If sheetName = chkSheetName Then
      result = True
      Exit For
    Else
      result = False
    End If
  Next
  chkExcludeSheet = result
End Function


'**************************************************************************************************
' * 配列が空かどうか
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
 Function chkArrayEmpty(arrayTmp As Variant) As Boolean

  On Error GoTo catchError

  If UBound(arrayTmp) >= 0 Then
    chkArrayEmpty = False
  Else
    chkArrayEmpty = True
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
  chkArrayEmpty = True

End Function


'**************************************************************************************************
' * ブックが開かれているかチェック
' *
' * @Link https://www.moug.net/tech/exvba/0060042.html
'**************************************************************************************************
Function chkBookOpened(chkFile) As Boolean
  Dim myChkBook As Workbook
  
  On Error Resume Next
  Set myChkBook = Workbooks(chkFile)

  If Err.Number > 0 Then
    chkBookOpened = False
  Else
    chkBookOpened = True
  End If
End Function


'**************************************************************************************************
' * ヘッダーチェック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHeader(baseNameArray As Variant, chkNameArray As Variant)
  Dim errMeg As String
  Dim i As Integer

  On Error GoTo catchError
  errMeg = ""

  If UBound(baseNameArray) <> UBound(chkNameArray) Then
    errMeg = "個数が異なります。"
    errMeg = errMeg & vbNewLine & UBound(baseNameArray) & "<=>" & UBound(chkNameArray) & vbNewLine
  Else
    For i = LBound(baseNameArray) To UBound(baseNameArray)
      If baseNameArray(i) <> chkNameArray(i) Then
        errMeg = errMeg & vbNewLine & i & ":" & baseNameArray(i) & "<=>" & chkNameArray(i)
      End If
    Next
  End If
  chkHeader = errMeg

  Exit Function
'エラー発生時------------------------------------
catchError:
  chkHeader = "エラーが発生しました"
End Function


'**************************************************************************************************
' * データチェック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'日付
Function chkIsDate(chkVal As Date, startDay As Date, endDay As Date)
  Dim chkFlg As Boolean
  chkFlg = False
  
  If IsDate(chkVal) = True Then
    If startDay <= chkVal And chkVal <= endDay Then
      chkFlg = True
    Else
      chkFlg = False
    End If
  Else
    chkFlg = False
  End If
  
  chkIsDate = chkFlg
End Function

'**************************************************************************************************
' * ファイルの保存場所がローカルディスクかどうか判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkLocalDrive(targetPath As String)
  Dim FSO As Object
  Dim driveName As String
  Dim driveType As Long
  Dim retVal As Boolean

  Set FSO = CreateObject("Scripting.FileSystemObject")
  driveName = FSO.GetDriveName(targetPath)
  
  'ドライブの種類を判別
  If driveName = "" Then
      driveType = 0 '不明
  Else
      driveType = FSO.GetDrive(driveName).driveType
  End If

  Select Case driveType
    Case 1
      retVal = True
      Call Library.showDebugForm("Library.chkLocalDrive", "リムーバブルディスク")
    Case 2
      retVal = True
      Call Library.showDebugForm("Library.chkLocalDrive", "ハードディスク")
    Case Else
      retVal = False
      Call Library.showDebugForm("Library.chkLocalDrive", "不明、ネットワークドライブ、CDドライブなど")
  End Select

  If BK_setVal("debugMode") = "develop" Then
    retVal = False
  End If
  chkLocalDrive = retVal

  Exit Function
'エラー発生時------------------------------------
catchError:
End Function


'**************************************************************************************************
' * パスからファイルかディレクトリかを判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkPathDecision(targetPath As String)
  Dim FSO As Object
  Dim retVal As String
  Dim targetType

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FolderExists(targetPath) Then
    retVal = "dir"
  Else
    If FSO.FileExists(targetPath) Then
      targetType = FSO.GetExtensionName(targetPath)
      retVal = UCase(targetType)
    End If
  End If
  Set FSO = Nothing
  
  chkPathDecision = retVal
End Function


'**************************************************************************************************
' * ファイルの存在確認
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkFileExists(targetPath As String)
  Dim FSO As Object

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FileExists(targetPath) Then
    chkFileExists = True
  Else
    chkFileExists = False
  End If
  Set FSO = Nothing

End Function


'**************************************************************************************************
' * ディレクトリの存在確認
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkDirExists(targetPath As String)
  Dim FSO As Object

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FolderExists(targetPath) Then
    chkDirExists = True
  Else
    chkDirExists = False
  End If
  Set FSO = Nothing
End Function


'**************************************************************************************************
' * ByteからKB,MB,GBへ変換
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convscale(ByVal lngVal As Long) As String
  Dim convVal As String

  If lngVal >= 1024 ^ 3 Then
    convVal = Round(lngVal / (1024 ^ 3), 3) & " GB"
  
  ElseIf lngVal >= 1024 ^ 2 Then
    convVal = Round(lngVal / (1024 ^ 2), 3) & " MB"
    
  ElseIf lngVal >= 1024 Then
    convVal = Round(lngVal / (1024), 3) & " KB"
  Else
    convVal = lngVal & " Byte"
  End If

  convscale = convVal
End Function


'**************************************************************************************************
' * 固定長文字列に変換
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFixedLength(strTarget As String, lengs As Long, addString As String) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  Do While LenB(strTarget) <= lengs
    strTarget = strTarget & addString
  Loop
  convFixedLength = strTarget
End Function


'**************************************************************************************************
' * キャメルケースをスネークケースに変換
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function covCamelToSnake(ByVal val As String, Optional ByVal isUpper As Boolean = False) As String
  Dim ret As String
  Dim i      As Long, Length As Long

  Length = Len(val)
  For i = 1 To Length
    If UCase(Mid(val, i, 1)) = Mid(val, i, 1) Then
      If i = 1 Then
        ret = ret & Mid(val, i, 1)
      ElseIf i > 1 And UCase(Mid(val, i - 1, 1)) = Mid(val, i - 1, 1) Then
        ret = ret & Mid(val, i, 1)
      Else
        ret = ret & "_" & Mid(val, i, 1)
      End If
    Else
      ret = ret & Mid(val, i, 1)
    End If
  Next

  If isUpper Then
    covCamelToSnake = UCase(ret)
  Else
    covCamelToSnake = LCase(ret)
  End If
End Function


'**************************************************************************************************
' * スネークケースをキャメルケースに変換
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function convSnakeToCamel(ByVal val As String, Optional ByVal isFirstUpper As Boolean = False) As String
  Dim ret As String
  Dim i   As Long
  Dim snakeSplit As Variant

  snakeSplit = Split(val, "_")

  For i = LBound(snakeSplit) To UBound(snakeSplit)
    ret = ret & UCase(Mid(snakeSplit(i), 1, 1)) & Mid(snakeSplit(i), 2, Len(snakeSplit(i)))
  Next

  If isFirstUpper Then
    convSnakeToCamel = ret
  Else
    convSnakeToCamel = LCase(Mid(ret, 1, 1)) & Mid(ret, 2, Len(ret))
  End If
End Function


'**************************************************************************************************
' * 半角のカタカナを全角のカタカナに変換する(ただし英数字は半角にする)
' *
' * @link   http://officetanaka.net/excel/function/tips/tips45.htm
'**************************************************************************************************
Function convHan2Zen(Text As String) As String
  Dim i As Long, buf As String
  Dim c As Range
  Dim rData As Variant, ansData As Variant

  For i = 1 To Len(Text)
    DoEvents
    rData = StrConv(Text, vbWide)
    If Mid(rData, i, 1) Like "[Ａ-ｚ]" Or Mid(rData, i, 1) Like "[０-９]" Or Mid(rData, i, 1) Like "[－！（）／]" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
      ansData = ansData & Mid(rData, i, 1)
    End If
  Next i
  convHan2Zen = ansData
End Function


'**************************************************************************************************
' * パイプをカンマに変換
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convPipe2Comma(strText As String) As String
  Dim covString As String
  Dim tmp As Variant
  Dim i As Integer
  
  tmp = Split(strText, "|")
  covString = ""
  For i = 0 To UBound(tmp)
    If i = 0 Then
      covString = tmp(i)
    Else
      covString = covString & "," & tmp(i)
    End If
  Next i
  convPipe2Comma = covString
End Function


'**************************************************************************************************
' * Base64エンコード(ファイル)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" '初期化
  On Error Resume Next
  Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
  With CreateObject("ADODB.Stream")
    .Type = adTypeBinary
    .Open
    .LoadFromFile filePath
    elm.dataType = "bin.base64"
    elm.nodeTypedValue = .Read(adReadAll)
    ret = elm.Text
    .Close
  End With
  On Error GoTo 0
  convBase64EncodeForFile = ret
End Function


'**************************************************************************************************
' * Base64エンコード(文字列)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String
  Dim ret As String
  Dim d() As Byte
  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" '初期化
  On Error Resume Next
  With CreateObject("ADODB.Stream")
    .Open
    .Type = adTypeText
    .Charset = "UTF-8"
    .WriteText str
    .Position = 0
    .Type = adTypeBinary
    .Position = 3
    d = .Read()
    .Close
  End With
  With CreateObject("MSXML2.DOMDocument").createElement("base64")
    .dataType = "bin.base64"
    .nodeTypedValue = d
    ret = .Text
  End With
  On Error GoTo 0
  convBase64EncodeForString = ret
End Function


'**************************************************************************************************
' * URL-safe Base64エンコード
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLSafeBase64Encode(ByVal str As String) As String

  str = convBase64EncodeForString(str)
  str = Replace(str, "+", "-")
  str = Replace(str, "/", "_")

  convURLSafeBase64Encode = str
End Function


'**************************************************************************************************
' * URLエンコード
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLEncode(ByVal str As String) As String
  Dim EncodeURL As String
  
  With CreateObject("ScriptControl")
    .Language = "JScript"
    EncodeURL = .codeobject.encodeURIComponent(str)
  End With
  
  convURLEncode = EncodeURL
End Function


'**************************************************************************************************
' * 先頭１文字目を大文字化
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFirstCharConvert(ByVal strTarget As String) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  strFirst = UCase(Left$(strTarget, 1))
  strExceptFirst = Mid$(strTarget, 2, Len(strTarget))
  convFirstCharConvert = strFirst & strExceptFirst
End Function


'**************************************************************************************************
' * 文字列の左側から指定文字数削除する関数
' *
' * @Link   https://vbabeginner.net/vbaで文字列の右側や左側から指定文字数削除する/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
  Dim iLen    As Long

  '文字列ではない場合
  If VarType(s) <> vbString Then
      cutLeft = s & "文字列ではない"
      Exit Function
  End If

  iLen = Len(s)

  '文字列長より指定文字数が大きい場合
  If iLen < i Then
      cutLeft = s & "文字列長より指定文字数が大きい"
      Exit Function
  End If

  cutLeft = Right(s, iLen - i)
End Function


'**************************************************************************************************
' * 文字列の右側から指定文字数削除する関数
' *
' * @Link   https://vbabeginner.net/vbaで文字列の右側や左側から指定文字数削除する/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
  Dim iLen    As Long

  If VarType(s) <> vbString Then
    cutRight = s & "文字列ではない"
    Exit Function
  End If

  iLen = Len(s)

  '文字列長より指定文字数が大きい場合
  If iLen < i Then
    cutRight = s & "文字列長より指定文字数が大きい"
    Exit Function
  End If

  cutRight = Left(s, iLen - i)
End Function


'**************************************************************************************************
' * 連続改行の削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delMultipleLine(targetValue As String)
  Dim combineMultipleLine As String
  
  With CreateObject("VBScript.RegExp")
    .Global = True
    .Pattern = "(\r\n)+"
    combineMultipleLine = .Replace(targetValue, vbCrLf)
  End With
  
  delMultipleLine = combineMultipleLine
End Function


'**************************************************************************************************
' * シート削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delSheetData(targetSheet As Worksheet, Optional line As Long)

  Call Library.showDebugForm("funcName", "Library.delSheetData")
  Call Library.showDebugForm("delLine  ", line)
  
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  Call Library.showDebugForm("sheetName", targetSheet.Name)
  
  If line <> 0 Then
    targetSheet.Rows(line & ":" & Rows.count).delete Shift:=xlUp
    targetSheet.Rows(line & ":" & Rows.count).Select
    targetSheet.Rows(line & ":" & Rows.count).NumberFormatLocal = "G/標準"
    targetSheet.Rows(line & ":" & Rows.count).style = "Normal"
  Else
    targetSheet.Cells.delete Shift:=xlUp
    targetSheet.Cells.NumberFormatLocal = "G/標準"
    targetSheet.Cells.style = "Normal"
  End If
  DoEvents

  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * セル内の改行削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delCellLinefeed(val As String)
  Dim stringVal As Variant
  Dim retVal As String
  Dim count As Integer

  retVal = ""
  count = 0
  For Each stringVal In Split(val, vbLf)
    If stringVal <> "" And count <= 1 Then
      retVal = retVal & stringVal & vbLf
      count = 0
    Else
      count = count + 1
    End If
  Next
  delCellLinefeed = retVal
End Function


'**************************************************************************************************
' * 選択範囲の画像削除
' *
' * @Link https://www.relief.jp/docs/018407.html
'**************************************************************************************************
Function delImage()
  Dim Rng As Range
  Dim shp As Shape

  If TypeName(Selection) <> "Range" Then
    Exit Function
  End If

  For Each shp In ActiveSheet.Shapes
    Set Rng = Range(shp.TopLeftCell, shp.BottomRightCell)

    If Not (Intersect(Rng, Selection) Is Nothing) Then
      shp.delete
    End If
  Next
End Function


'**************************************************************************************************
' * セルの名称設定削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delVisibleNames()
  Dim Name As Object

  On Error Resume Next

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.delete
    End If
  Next
End Function


'**************************************************************************************************
' * テーブルデータ削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delTableData()
  Dim endLine As Long

  On Error Resume Next

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  Rows("3:" & endLine).Select
  Selection.delete Shift:=xlUp

  Rows("2:3").Select
  Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents

  Cells.Select
  Selection.NumberFormatLocal = "G/標準"

  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * ファイルコピー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execCopy(srcPath As String, dstPath As String)
  Dim FSO As Object
  
  On Error GoTo catchError
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call showDebugForm("  コピー元：" & srcPath)
  Call showDebugForm("  コピー先：" & dstPath)
  
  If chkFileExists(srcPath) = False Then
    Call showNotice(404, "コピー元", True)
  End If
  
  If chkDirExists(getParentDir(dstPath)) = False Then
    Call Library.execMkdir(getParentDir(dstPath))
  End If
  FSO.CopyFile srcPath, dstPath
  Set FSO = Nothing

  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'**************************************************************************************************
' * ファイル移動
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execMove(srcPath As String, dstPath As String)
  Dim FSO As Object
  
  On Error GoTo catchError
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call showDebugForm("移動元：" & srcPath)
  Call showDebugForm("移動先：" & dstPath)
  
  If chkFileExists(srcPath) = False Then
    Call showNotice(404, "移動元", True)
  End If
  FSO.MoveFile srcPath, dstPath
  Set FSO = Nothing

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'**************************************************************************************************
' * ファイル削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execDel(srcPath As String)
  Dim FSO As Object
  
  On Error GoTo catchError
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call showDebugForm("  削除対象：" & srcPath)
  
  If srcPath Like "*[*]*" Then
  
  ElseIf chkFileExists(srcPath) = False Then
    Call showNotice(404, "削除対象", True)
  End If
  
  FSO.DeleteFile srcPath
  Set FSO = Nothing

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

'**************************************************************************************************
' * ファイル名変更
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execRename(srcPath As String, oldFileName As String, fileName As String, Optional errMeg As String)
  Dim FSO As Object
  Dim errFlg As Boolean
  Const funcName As String = "Library.execReName"

  On Error GoTo catchError

  errFlg = False
  Call showDebugForm("変更元", srcPath)
  Call showDebugForm("旧名称", oldFileName)
  Call showDebugForm("新名称", fileName)
  
  If chkFileExists(srcPath & "\" & oldFileName) = False Then
    If IsMissing(errMeg) Then
      Call showNotice(404, "変更元", True)
    Else
      errMeg = "変更元のファイルがありません[" & oldFileName & "]"
      errFlg = True
    End If
  End If
  If chkFileExists(srcPath & "\" & fileName) = True Then
    If IsMissing(errMeg) Then
      Call showNotice(414, fileName, True)
    Else
      errMeg = "同名のファイルが存在します[" & fileName & "]"
      errFlg = True
    End If
  End If
  If errFlg = False Then
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.GetFile(srcPath & "\" & oldFileName).Name = fileName
    Set FSO = Nothing
    
    execRename = True
  Else
    execRename = False
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, "[" & Err.Number & "] " & Err.Description)
  errMeg = Err.Description
  
  execRename = False
End Function

'**************************************************************************************************
' * MkDirで階層の深いフォルダーを作る
' *
' * @link http://officetanaka.net/excel/vba/filesystemobject/sample10.htm
'**************************************************************************************************
Function execMkdir(fullPath As String)
  
  If chkDirExists(fullPath) Then
    Exit Function
  End If
  Call chkParentDir(fullPath)
End Function

'==================================================================================================
Private Function chkParentDir(TargetFolder)
  Dim ParentFolder As String, FSO As Object

  On Error GoTo catchError
  Set FSO = CreateObject("Scripting.FileSystemObject")

  ParentFolder = FSO.GetParentFolderName(TargetFolder)
  If Not FSO.FolderExists(ParentFolder) Then
    Call chkParentDir(ParentFolder)
  End If

  FSO.CreateFolder TargetFolder
  Set FSO = Nothing

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "ディレクトリの作成に失敗しました" & vbNewLine & Err.Description, True)
End Function


'**************************************************************************************************
' * zip圧縮/解凍
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execCompress(srcPath As String, zipFilePath As String) As Boolean
  Dim sh  As New IWshRuntimeLibrary.WshShell
  Dim ex  As WshExec
  Dim cmd As String
  
  Call showDebugForm("対象ディレクトリ：" & srcPath)
  Call showDebugForm("zipファイル     ：" & zipFilePath)
  
  If chkDirExists(srcPath) = False Then
    Call showNotice(403, "対象ディレクトリ", True)
  End If
  
  '// 半角スペースをバッククォートでエスケープ
  srcPath = Replace(srcPath, " ", "` ")
  zipFilePath = Replace(zipFilePath, " ", "` ")
  
  cmd = "Compress-Archive -Path " & srcPath & " -DestinationPath " & zipFilePath & " -Force"
  Call showDebugForm("cmd：" & cmd)
  Set ex = sh.exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & cmd)
  
  If ex.Status = WshFailed Then
    execCompress = False
    Exit Function
  End If
  
  Do While ex.Status = WshRunning
    DoEvents
  Loop
  
  execCompress = True
End Function


'==================================================================================================
Function execUncompress(zipFilePath As String, dstPath As String) As Boolean
  Dim sh As New IWshRuntimeLibrary.WshShell
  Dim ex As WshExec
  Dim cmd As String
  
  Call showDebugForm("zipファイル　　 ：" & zipFilePath)
  Call showDebugForm("対象ディレクトリ：" & dstPath)
  
  If chkFileExists(zipFilePath) = False Then
    Call showNotice(404, "解凍対象", True)
  End If
  If chkDirExists(dstPath) = False Then
    Call showNotice(403, "解凍先", True)
  End If
  
  '// 半角スペースをバッククォートでエスケープ
  zipFilePath = Replace(zipFilePath, " ", "` ")
  dstPath = Replace(dstPath, " ", "` ")
  
  cmd = "Expand-Archive -Path " & zipFilePath & " -DestinationPath " & dstPath & " -Force"
  Call showDebugForm("cmd：" & cmd)
  Set ex = sh.exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & cmd)
  
  If ex.Status = WshFailed Then
    execUncompress = False
    Exit Function
  End If
  Do While ex.Status = WshRunning
    DoEvents
  Loop
  execUncompress = True
End Function


'**************************************************************************************************
' * PC、Office等の情報取得
' * 連想配列を利用しているので、Microsoft Scripting Runtimeが必須
' * MachineInfo.Item ("Excel") で呼び出し
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object

  On Error Resume Next
  
  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  ' OSのバージョン取得-----------------------------------------------------------------------------
  Select Case Application.OperatingSystem

    Case "Windows (64-bit) NT 6.01"
        MachineInfo.add "OS", "Windows7-64"

    Case "Windows (32-bit) NT 6.01"
        MachineInfo.add "OS", "Windows7-32"

    Case "Windows (32-bit) NT 5.01"
        MachineInfo.add "OS", "WindowsXP-32"

    Case "Windows (64-bit) NT 5.01"
        MachineInfo.add "OS", "WindowsXP-64"

    Case Else
       MachineInfo.add "OS", Application.OperatingSystem
  End Select

  ' Excelのバージョン取得--------------------------------------------------------------------------
  Select Case Application.Version
    Case "16.0"
        MachineInfo.add "Excel", "2016"
    Case "14.0"
        MachineInfo.add "Excel", "2010"
    Case "12.0"
        MachineInfo.add "Excel", "2007"
    Case "11.0"
        MachineInfo.add "Excel", "2003"
    Case "10.0"
        MachineInfo.add "Excel", "2002"
    Case "9.0"
        MachineInfo.add "Excel", "2000"
    Case Else
       MachineInfo.add "Excel", Application.Version
  End Select

  'PCの情報----------------------------------------------------------------------------------------
  MachineInfo.add "UserName", WshNetworkObject.UserName
  MachineInfo.add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.add "UserDomain", WshNetworkObject.UserDomain

  '画面の解像度等取得------------------------------------------------------------------------------
  MachineInfo.add "monitors", GetSystemMetrics(80)
  MachineInfo.add "displayX", GetSystemMetrics(0)
  MachineInfo.add "displayY", GetSystemMetrics(1)
  
  MachineInfo.add "displayVirtualX", GetSystemMetrics(78)
  MachineInfo.add "displayVirtualY", GetSystemMetrics(79)
  MachineInfo.add "appTop", ActiveWindow.Top
  MachineInfo.add "appLeft", ActiveWindow.Left
  MachineInfo.add "appWidth", ActiveWindow.Width
  MachineInfo.add "appHeight", ActiveWindow.Height
  
  Set WshNetworkObject = Nothing
End Function


'**************************************************************************************************
' * 文字数カウント
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getByteString(arryColumn As String, Optional line As Long) As Long
  Dim colLineName As Variant
  Dim count As Long

  count = 0
  For Each colLineName In Split(arryColumn, ",")
    If line > 0 Then
      count = count + LenB(Range(colLineName & line).Value)
    Else
      count = count + LenB(Range(colLineName).Value)
    End If
  Next colLineName

  getByteString = count
End Function


'**************************************************************************************************
' * セルの座標取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getCellPosition(Rng As Range, ActvCellTop As Long, ActvCellLeft As Long)
  Dim R1C1Top As Long, R1C1Left As Long
  Dim DPI, PPI
'  Const DPI As Long = 96
'  Const PPI As Long = 72
  
  R1C1Top = ActiveWindow.PointsToScreenPixelsY(0)
  R1C1Left = ActiveWindow.PointsToScreenPixelsX(0)

'  ActvCellTop = ((R1C1Top * DPI / PPI) * (ActiveWindow.Zoom / 100)) + Rng.Top
'  ActvCellLeft = ((R1C1Left * DPI / PPI) * (ActiveWindow.Zoom / 100)) + Rng.Left

  ActvCellTop = (((Rng.Top * (DPI / PPI)) * (ActiveWindow.Zoom / 100)) + R1C1Top) * (PPI / DPI)
  ActvCellLeft = (((Rng.Left * (DPI / PPI)) * (ActiveWindow.Zoom / 100)) + R1C1Left) * (PPI / DPI)

'  If ActvCellLeft <= 0 Then
'    ActvCellLeft = 20
'  End If

  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("R1C1Top ：" & R1C1Top)
  Call Library.showDebugForm("R1C1Left ：" & R1C1Left)
  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("Rng.Address ：" & Rng.Address)
  Call Library.showDebugForm("ActvCellTop ：" & ActvCellTop)
  Call Library.showDebugForm("ActvCellLeft：" & ActvCellLeft)
End Function


'**************************************************************************************************
' * 列名から列番号を求める
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long
  getColumnNo = Range(targetCell, targetCell).Column
End Function


'**************************************************************************************************
' * 列番号から列名を求める
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String
  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function


'**************************************************************************************************
' * カラーパレットを表示し、色コードを取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getColor(colorValue As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Call getRGB(colorValue, Red, Green, Blue)
  Application.Dialogs(xlDialogEditColor).Show 10, Red, Green, Blue

  setColorValue = ActiveWorkbook.Colors(10)
  If setColorValue = False Then
    setColorValue = colorValue
  End If

  getColor = setColorValue
End Function


'**************************************************************************************************
' * フォントダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFont(FontName As String, fontSize As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Application.Dialogs(xlDialogActiveCellFont).Show FontName, "レギュラー", fontSize
End Function


'**************************************************************************************************
' * IndentLevel値取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getIndentLevel(targetRange As Range)
  Dim thisTargetSheet As Worksheet

  Application.Volatile

  If targetRange = "" Then
    getIndentLevel = ""
  Else
    getIndentLevel = targetRange.IndentLevel + 1
  End If
End Function


'**************************************************************************************************
' * RGB値取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function


'**************************************************************************************************
' * ディレクトリ選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getDirPath(CurrentDirectory As String, Optional title As String)

  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = CurrentDirectory & "\"
    .AllowMultiSelect = False

    If title <> "" Then
      .title = title & "の場所を選択してください"
    Else
      .title = "フォルダーを選択してください"
    End If

    If .Show = True Then
      getDirPath = .SelectedItems(1)
    Else
      getDirPath = ""
    End If
  End With
End Function


'**************************************************************************************************
' * ファイル保存ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)
  Dim filePath As String
  Dim result As Long
  Dim fileName As Variant

  fileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm,すべてのファイル, *.*", _
      FilterIndex:=FileTypeNo)

  If fileName <> "False" Then
    getSaveFilePath = filePath
  Else
    getSaveFilePath = ""
  End If
End Function

'**************************************************************************************************
' * ファイル選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, fileName As String, title As String, fileType As String)

  Dim filePath As String
  Dim result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' ファイルの種類を設定
    .Filters.Clear
    Select Case fileType
      Case "Excel"
            .Filters.add "Excelブック", "*.xls; *.xlsx; *.xlsm"
      Case "txt"
        .Filters.add "テキストファイル", "*.txt"
      
      Case "csv"
        .Filters.add "CSVファイル", "*.csv"
      
      Case "json"
        .Filters.add "JSONファイル", "*.json"
      
      Case "sql"
        .Filters.add "SQLファイル", "*.sql"
      
      Case "mdb"
        .Filters.add "Accesssデータベース", "*.mdb;*.accdb"
      
      Case "img"
        .Filters.add "イメージファイル", "*.bmp;*.jpg;*.gif;*.png"
      
      Case "psd"
        .Filters.add "Photoshop Data", "*.psd"
      
      Case "クリエイティブ"
        .Filters.add "クリエイティブ", "*.jpg;*.gif;*.png;*.mp4"
      
      Case "mov"
        .Filters.add "動画ファイル", "*.mp4"

      
      Case Else
        .Filters.add "すべてのファイル", "*.*"
    End Select
    '.FilterIndex = FileTypeNo

    '表示するフォルダ
    If chkDirExists(CurrentDirectory) = True Then
      .InitialFileName = CurrentDirectory & "\" & fileName
    Else
      .InitialFileName = ActiveWorkbook.Path & "\" & fileName
    End If
    
    '表示形式の設定
    .InitialView = msoFileDialogViewWebView

    'ダイアログ ボックスのタイトル設定
    .title = title & "を選択してください"


    If .Show = -1 Then
      filePath = .SelectedItems(1)
    Else
      filePath = ""
    End If
  End With

  getFilePath = filePath

End Function


'**************************************************************************************************
' * 複数ファイル選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, fileName As String, title As String, fileType As String)

  Dim filePath() As Variant
  Dim result As Long
  Dim i As Integer

  With Application.FileDialog(msoFileDialogFilePicker)
    '複数選択を許可
    .AllowMultiSelect = True

    ' ファイルの種類を設定
    .Filters.Clear
    Select Case fileType
      Case "Excel"
            .Filters.add "Excelブック", "*.xls; *.xlsx; *.xlsm"
      Case "txt"
        .Filters.add "テキストファイル", "*.txt"
      
      Case "csv"
        .Filters.add "CSVファイル", "*.csv"
      
      Case "json"
        .Filters.add "JSONファイル", "*.json"
      
      Case "sql"
        .Filters.add "SQLファイル", "*.sql"
      
      Case "mdb"
        .Filters.add "Accesssデータベース", "*.mdb;*.accdb"
      
      Case "img"
        .Filters.add "イメージファイル", "*.bmp;*.jpg;*.gif;*.png"
      
      Case "psd"
        .Filters.add "Photoshop Data", "*.psd"
      
      Case "クリエイティブ"
        .Filters.add "クリエイティブ", "*.jpg;*.gif;*.png;*.mp4"
      
      Case "mov"
        .Filters.add "動画ファイル", "*.mp4"
        
      Case Else
        .Filters.add "すべてのファイル", "*.*"
    End Select
    '.FilterIndex = FileTypeNo

    '表示するフォルダ
    .InitialFileName = CurrentDirectory & "\" & fileName

    '表示形式の設定
    .InitialView = msoFileDialogViewWebView

    'ダイアログ ボックスのタイトル設定
    .title = title

    If .Show = -1 Then
      ReDim Preserve filePath(.SelectedItems.count - 1)
      For i = 1 To .SelectedItems.count
        filePath(i - 1) = .SelectedItems(i)
      Next i
    Else
      ReDim Preserve filePath(0)
      filePath(0) = ""
    End If
  End With

  getFilesPath = filePath

End Function

'**************************************************************************************************
' * ディレクトリ内のファイル一覧取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileList(Path As String, fileName As String)
  Dim f As Object, cnt As Long
  Dim list() As String

  cnt = 0
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(Path).Files
      If f.Name Like fileName Then
        ReDim Preserve list(cnt)
        list(cnt) = f.Name
        cnt = cnt + 1
      End If
    Next f
  End With

  getFileList = list
End Function


'**************************************************************************************************
' * ファイル情報取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileInfo(targetFilePath As String, Optional fileInfo As Object, Optional getType As String)
  Dim FSO As Object
  Dim fileObject As Object
  Dim sp As Shape
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Set fileInfo = Nothing
  Set fileInfo = CreateObject("Scripting.Dictionary")
  
  '作成日時
  fileInfo.add "createAt", Format(FSO.GetFile(targetFilePath).DateCreated, "yyyy/mm/dd hh:nn:ss")
  
  '更新日時
  fileInfo.add "updateAt", Format(FSO.GetFile(targetFilePath).DateLastModified, "yyyy/mm/dd hh:nn:ss")
  
  'ファイルサイズ
  fileInfo.add "size", FSO.GetFile(targetFilePath).Size
  
  'ファイルの種類
  fileInfo.add "type", FSO.GetFile(targetFilePath).Type
  
  '拡張子
  fileInfo.add "extension", FSO.GetExtensionName(targetFilePath)
  
  'ファイル名
  fileInfo.add "fileName", FSO.GetFile(targetFilePath).Name
  
  'ファイルが存在するフォルダ
  fileInfo.add "CurrentDir", FSO.GetFile(targetFilePath).ParentFolder
  
  Select Case FSO.GetExtensionName(targetFilePath)
    Case "mp4"
    
    Case "png"
      Set sp = ActiveSheet.Shapes.AddPicture( _
                fileName:=targetFilePath, _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0, _
                Width:=0, _
                Height:=0 _
                )
      With sp
        .LockAspectRatio = msoTrue
        .ScaleHeight 1, msoTrue
        .ScaleWidth 1, msoTrue
        
        fileInfo.add "width", CLng(.Width * 4 / 3)
        fileInfo.add "height", CLng(.Height * 4 / 3)
        .delete
      End With
            
    Case "bmp", "jpg", "jpeg", "gif", "emf", "ico", "rle", "wmf"
      Set fileObject = LoadPicture(targetFilePath)
      fileInfo.add "width", fileObject.Width
      fileInfo.add "height", fileObject.Height
      
      Set fileObject = Nothing
    
    Case Else
  End Select
  
  Set FSO = Nothing

  If getType <> "" Then
    getFileInfo = fileInfo(getType)
    Set fileInfo = Nothing
  End If
  
End Function

'**************************************************************************************************
' * ファイルの親フォルダ取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getParentDir(targetPath As String) As String
  Dim parentDir As String
  
  parentDir = Left(targetPath, InStrRev(targetPath, "\") - 1)
'  Call Library.showDebugForm(" parentDir：" & parentDir)
  
  getParentDir = parentDir
End Function


'**************************************************************************************************
' * 指定バイト数の固定長データ作成(文字列処理)
' *
' * @Link http://www.asahi-net.or.jp/~ef2o-inue/vba_o/function05_110_055.html
'**************************************************************************************************
Function getFixlng(strInText As String, lngFixBytes As Long) As String
    Dim lngKeta As Long
    Dim lngByte As Long, lngByte2 As Long, lngByte3 As Long
    Dim ix As Long
    Dim intCHAR As Long
    Dim strOutText As String

    lngKeta = Len(strInText)
    strOutText = strInText
    ' バイト数判定
    For ix = 1 To lngKeta
        ' 1文字ずつ半角/全角を判断
        intCHAR = Asc(Mid(strInText, ix, 1))
        ' 全角と判断される場合はバイト数に1を加える
        If ((intCHAR < 0) Or (intCHAR > 255)) Then
            lngByte2 = 2        ' 全角
        Else
            lngByte2 = 1        ' 半角
        End If
        ' 桁あふれ判定(右切り捨て)
        lngByte3 = lngByte + lngByte2
        If lngByte3 >= lngFixBytes Then
            If lngByte3 > lngFixBytes Then
                strOutText = Left(strInText, ix - 1)
            Else
                strOutText = Left(strInText, ix)
                lngByte = lngByte3
            End If
            Exit For
        End If
        lngByte = lngByte3
    Next ix
    ' 桁不足判定(空白文字追加)
    If lngByte < lngFixBytes Then
        strOutText = strOutText & Space(lngFixBytes - lngByte)
    End If
    getFixlng = strOutText
End Function


'**************************************************************************************************
' * 選択セルの拡大表示呼出
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionForm(Text As String, SetSelectTargetRows As String)
  With Frm_Zoom
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 10)
    .Left = Application.Left + (ActiveWindow.Height / 5)
    .TextBox = Text
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    .Caption = SetSelectTargetRows
    
    .Show vbModeless
  End With
End Function


'**************************************************************************************************
' * デバッグ用画面表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(ByVal meg1 As String, Optional meg2 As Variant)
  Dim runTime As Date
  Dim StartUpPosition As Long
  Const funcName As String = "Library.showDebugForm"
  
  On Error GoTo catchError

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  If BK_setVal("debugMode") = "none" Then
    Exit Function
  End If

  meg1 = Replace(meg1, vbNewLine, " ")
  If IsMissing(meg2) = False Then
    meg1 = "  " & meg1 & "：" & Application.WorksheetFunction.Trim(Replace(CStr(meg2), vbNewLine, " "))
  End If
  
  Select Case BK_setVal("debugMode")
    Case "file"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If

    Case "form"

    Case "all"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If

    Case "develop"
      If meg1 <> "" Then
        Debug.Print runTime & vbTab & meg1
        Call outputLog(runTime, meg1)
      End If

    Case Else
      Exit Function
  End Select

  DoEvents
  Exit Function

'エラー発生時=====================================================================================
catchError:
  Debug.Print (funcName & " [" & Err.Number & "]" & Err.Description)
End Function


'**************************************************************************************************
' * 処理情報通知
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional ReplaceMeg As String, Optional runEndflg As Boolean)
  Dim Message As String, SpeakMeg As String
  Dim runTime As Date
  Dim endLine As Long

  On Error GoTo catchError
  
  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  endLine = sheetNotice.Cells(Rows.count, 1).End(xlUp).Row
  Message = Application.WorksheetFunction.VLookup(Code, sheetNotice.Range("A2:B" & endLine), 2, False)
  Message = Replace(Message, "%%", ReplaceMeg)
  
  If ReplaceMeg <> "" Then
    Message = Replace(Message, "<>", ReplaceMeg)
  End If
  
  If Message <> "" Then
    Message = Replace(Message, "<BR>", vbNewLine)
  End If
  
  SpeakMeg = Message
  If endTime <> 0 Then
    Message = Message & " 処理時間：" & endTime
  End If
  
  If runEndflg = True Then
    SpeakMeg = SpeakMeg & " 処理を中止します"
  End If

  
  If setVal("debugMode") = "speak" Or setVal("debugMode") = "develop" Or setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=SpeakMeg, SpeakAsync:=True, SpeakXML:=True
  End If
  Message = Replace(Message, "<", vbNewLine)
  Message = Replace(Message, ">", "")

  Select Case Code
    Case 0 To 399
      Call MsgBox(Message, vbInformation, thisAppName)

    Case 400 To 499
      Call MsgBox(Message, vbCritical, thisAppName)

    Case 500 To 599
      Call MsgBox(Message, vbExclamation, thisAppName)

    Case 999

    Case Else
      Call MsgBox(Message, vbCritical, thisAppName)
  End Select

  Message = "[" & Code & "]" & Message
  Call Library.showDebugForm(Message)
  
  '画面描写制御終了処理
  If runEndflg = True Then
    Call endScript
    Call Ctl_ProgressBar.showEnd
    End
  End If

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call MsgBox(Message, vbCritical, thisAppName)

End Function


'**************************************************************************************************
' * ランダム
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setString As String, ByVal setStringCnt As Integer) As String
  Dim i, n
  Dim str1 As String

  For i = 1 To setStringCnt
    '乱数ジェネレータを初期化
    Randomize
    n = Int((Len(setString) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(setString, n, 1)
  Next i
  makeRandomString = str1
End Function

'==================================================================================================
Function makeRandomNo(minNo As Long, maxNo As Long) As String
  Randomize
  makeRandomNo = Int((maxNo - minNo + 1) * Rnd + minNo)
End Function

'==================================================================================================
Function makeRandomDigits(maxCount As Long) As String
  Dim makeVal As String
  Dim tmpVal As String
  Dim count As Integer
  
  For count = 1 To maxCount
    Randomize
    tmpVal = CStr(Int(10 * Rnd))
    
    If count = 1 And tmpVal = 0 Then
      tmpVal = 1
    End If
    makeVal = makeVal & tmpVal
  Next
  makeRandomDigits = makeVal
End Function


'**************************************************************************************************
' * ログ出力
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function outputLog(runTime As Date, Message As String)
  Dim fileTimestamp As Date

  On Error GoTo catchError
  
  If logFile = "" Then
    Debug.Print "ログファイルが設定されていません"
    End
    
  ElseIf chkFileExists(logFile) Then
    fileTimestamp = FileDateTime(logFile)
  Else
    fileTimestamp = DateAdd("d", -1, Date)
  End If

  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    If Format(Date, "yyyymmdd") = Format(fileTimestamp, "yyyymmdd") Then
      .LoadFromFile logFile
      .Position = .Size
    End If
    .WriteText runTime & vbTab & Message, 1
    .SaveToFile logFile, 2
    .Close
  End With
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Debug.Print "ログ出力失敗" & Err.Number & "：" & Err.Description
  Debug.Print logFile
End Function


'==================================================================================================
Function outputText(Message As String, outputFilePath)
  Open outputFilePath For Output As #1
  Print #1, Message
  Close #1
End Function


'**************************************************************************************************
' * CSVインポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @link   https://www.tipsfound.com/vba/18014
'**************************************************************************************************
Function importCsv(filePath As String, Optional readLine As Long, Optional TextFormat As Variant, Optional charCode As String = "sjis")
  Dim ws As Worksheet
  Dim qt As QueryTable
  Dim count As Long, line As Long, endLine As Long

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 1 Then
    endLine = 1
  Else
    endLine = endLine + 1
  End If

  If readLine < 1 Then
    readLine = 1
  End If

  Set ws = ActiveSheet
  Set qt = ws.QueryTables.add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A" & endLine))
  With qt
    If charCode = "sjis" Then
      .TextFilePlatform = 932          ' Shift-JIS を開く
    Else
      .TextFilePlatform = 65001        ' UTF-8
    End If
    .TextFileParseType = xlDelimited ' 文字で区切った形式
    .TextFileCommaDelimiter = True   ' 区切り文字はカンマ
    .TextFileStartRow = readLine     ' 1 行目から読み込み
    .AdjustColumnWidth = False       ' 列幅を自動調整しない
    .RefreshStyle = xlOverwriteCells '上書きを指定
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' 引用符の指定

    If IsArray(TextFormat) Then
      .TextFileColumnDataTypes = TextFormat
    End If

    .Refresh
    DoEvents
    .delete
  End With
  Set qt = Nothing
  Set ws = Nothing

  Call Library.startScript
End Function


'**************************************************************************************************
' * Excelファイルのインポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function importXlsx(filePath As String, targeSheet As String, targeArea As String, dictSheet As Worksheet, Optional passWord As String)

  On Error GoTo catchError
  If passWord <> "" Then
    Workbooks.Open fileName:=filePath, ReadOnly:=True, passWord:=passWord
  Else
    Workbooks.Open fileName:=filePath, ReadOnly:=True
  End If

  If Worksheets(targeSheet).Visible = False Then
    Worksheets(targeSheet).Visible = True
  End If
  Sheets(targeSheet).Select

  ActiveWorkbook.Sheets(targeSheet).Rows.Hidden = False
  ActiveWorkbook.Sheets(targeSheet).Columns.Hidden = False

  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

  ActiveWorkbook.Sheets(targeSheet).Range(targeArea).Copy
  dictSheet.Range("A1").PasteSpecial xlPasteValues

  Application.CutCopyMode = False
  ActiveWorkbook.Close SaveChanges:=False
  dictSheet.Range("A1").Select
  
  DoEvents
  Call Library.startScript

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'**************************************************************************************************
' * パスワード生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String
  Dim i As Integer
  Dim n
  
  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
    '乱数ジェネレータを初期化
    Randomize
    n = Int((Len(halfChar) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(halfChar, n, 1)
  Next i
  makePasswd = str1
End Function


'**************************************************************************************************
' * ハイライト化
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setHighLight(SetArea As String, DisType As Boolean, SetColor As String)

  Range(SetArea).Select

  '条件付き書式をクリア
  Selection.FormatConditions.delete

  If DisType = False Then
    '行だけ設定
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    '行と列に設定
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
  End If

  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = SetColor
'    .Interior.TintAndShade = 0
'    .Font.ColorIndex = 1
  End With
  Selection.FormatConditions(1).StopIfTrue = False


End Function

'==================================================================================================
Function unsetHighLight()
  Static xRow
  Static xColumn
  Dim pRow, pColumn
  
  pRow = Selection.Row
  pColumn = Selection.Column
  xRow = pRow
  xColumn = pColumn
  If xColumn <> "" Then
    With Columns(xColumn).Interior
      .ColorIndex = xlNone
    End With
    With Rows(xRow).Interior
      .ColorIndex = xlNone
    End With
  End If
End Function


'**************************************************************************************************
' * 文字列分割
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function splitString(targetString As String, separator As String, count As Integer) As String
  Dim tmp As Variant

  If targetString <> "" Then
    tmp = Split(targetString, separator)
    splitString = tmp(count)
  Else
    splitString = ""
  End If
End Function


'**************************************************************************************************
' * 配列の最後に追加する
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setArrayPush(arrName As Variant, str As Variant)
  Dim i As Long

  i = UBound(arrName)
  If i = 0 Then

  Else
    i = i + 1
    ReDim Preserve arrName(i)
  End If
  arrName(i) = str

End Function


'**************************************************************************************************
' * フォントカラー設定
' *
' * @Link https://vbabeginner.net/vbaでセルの指定文字列の色や太さを変更する/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font     'Fontオブジェクト
  Dim i               '引数文字列のセルの位置
  Dim iLen            '引数文字列の文字数
  Dim R   As Range    'セル範囲の１セル

  iLen = Len(a_sSearch)
  i = 1

  For Each R In Selection
    Do
      i = InStr(i, R.Value, a_sSearch)
      If (i = 0) Then
        i = 1
        Exit Do
      End If
      Set f = R.Characters(i, iLen).Font
      f.Color = a_lColor
      f.Bold = a_bBold
      i = i + 1
    Loop
  Next
End Function


'**************************************************************************************************
' * レジストリ関連
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setRegistry(RegistrySubKey As String, RegistryKey As String, setVal As Variant)

  If getRegistry(RegistrySubKey, RegistryKey) <> setVal And RegistryKey <> "" Then
    Call SaveSetting(thisAppName, RegistrySubKey, RegistryKey, setVal)
  End If
End Function

'==================================================================================================
Function getRegistry(RegistrySubKey As String, RegistryKey As String)
  Dim regVal As String

  On Error GoTo catchError

  If RegistryKey <> "" Then
    regVal = GetSetting(thisAppName, RegistrySubKey, RegistryKey)
  End If
  If regVal = "" Then
    getRegistry = 0
  Else
    getRegistry = regVal
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description)
End Function


'==================================================================================================
Function delRegistry(RegistrySubKey As String, Optional RegistryKey As String)
  Dim regVal As String

  On Error GoTo catchError
  If RegistryKey = "" Then
    Call DeleteSetting(thisAppName, RegistrySubKey)
  Else
    Call DeleteSetting(thisAppName, RegistrySubKey, RegistryKey)
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
'  Call Library.showNotice(400, Err.Description, True)
End Function


'**************************************************************************************************
' * 参照設定を自動で行う
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setReferences(BookType As String)
'
'  On Error GoTo Err_SetReferences:
'
'  'Microsoft Scripting Runtime (Windows Script Host / FileSystemObject)----------------------------
'    LibScript = "C:\Windows\System32\scrrun.dll"
'    If Dir(LibScript) <> "" Then
'      ActiveWorkbook.VBProject.References.AddFromFile (LibScript)
'    Else
'      MsgBox ("Microsoft Scripting Runtimeを利用できません。" & vbLf & "利用できない機能があります")
'    End If
'
'  'Microsoft ActiveX Data Objects Library 6.1 (ADO)------------------------------------------------
'  If BookType = "DataBase" Then
'    LibADO = "C:\Program Files\Common Files\System\Ado\msado15.dll"
'    If Dir(LibADO) <> "" Then
'      ActiveWorkbook.VBProject.References.AddFromFile (LibADO)
'    Else
'      MsgBox ("Microsoft ActiveX Data Objectsを利用できません" & vbLf & "利用できない機能があります")
'    End If
'
'  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
'  LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
'    If Dir(LibDAO) <> "" Then
'      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
'    Else
'      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
'      If Dir(LibDAO) <> "" Then
'        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
'      Else
'        MsgBox ("Microsoft DAO 3.6 Objects Libraryを利用できません" & vbLf & "DBへの接続機能が利用できません")
'      End If
'    End If
'  End If
'
'  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
'  If BookType = "" Then
'    LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
'    If Dir(LibDAO) <> "" Then
'      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
'    Else
'      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
'      If Dir(LibDAO) <> "" Then
'        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
'      Else
'        MsgBox ("Microsoft DAO 3.6 Objects Libraryを利用できません" & vbLf & "DBへの接続機能が利用できません")
'      End If
'    End If
'  End If
'
'
'Func_Exit:
'  Set Ref = Nothing
'  Exit Function
'
'Err_SetReferences:
'  If Err.Number = 32813 Then
'    Resume Next
'  ElseIf Err.Number = 1004 Then
'    MsgBox ("「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」に変更を！")
'  Else
'    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
'    GoTo Func_Exit:
'  End If
End Function


'**************************************************************************************************
' * 選択セルの行背景設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor(SetArea As String, DisType As Boolean, SetColor As String)

  Range(SetArea).Select

  '条件付き書式をクリア
  Selection.FormatConditions.delete

  If DisType = False Then
    '行だけ設定
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    '行と列に設定
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
  End If

  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = SetColor
'    .Interior.TintAndShade = 0
'    .Font.ColorIndex = 1
  End With
  Selection.FormatConditions(1).StopIfTrue = False
End Function


'**************************************************************************************************
' * シートの保護/保護解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setProtectSheet(Optional thisAppPasswd As String)

  ActiveSheet.Protect passWord:=thisAppPasswd, DrawingObjects:=True, Contents:=True, Scenarios:=True
  ActiveSheet.EnableSelection = xlUnlockedCells

End Function

'==================================================================================================
Function unsetProtectSheet(Optional thisAppPasswd As String)

  ActiveSheet.Unprotect passWord:=thisAppPasswd
End Function


'**************************************************************************************************
' * 最初のシートを選択
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setFirstsheet()
  Dim i As Long

  For i = 1 To Sheets.count
    If Sheets(i).Visible = xlSheetVisible Then
      Sheets(i).Select
      Exit Function
    End If
  Next i
End Function


'**************************************************************************************************
' * 値の設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setValandRange(keyName As String, val As String)
  Range(keyName) = val

  If BK_setVal Is Nothing Then
    Call init.setting
  Else
    BK_setVal(keyName) = val
  End If
End Function


'**************************************************************************************************
' * バッチファイル実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function runBat(fileName As String)
  Dim obj As WshShell
  Dim rtnVal As String
  
  Set obj = New WshShell
  rtnVal = obj.run(fileName, WaitOnReturn:=True)

  Call Library.showDebugForm("実行ファイル", fileName)
  Call Library.showDebugForm("戻り値      ", rtnVal)

  runBat = rtnVal
End Function
    
    
'**************************************************************************************************
' * ファイル全体の文字列置換
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(fileName As String, TargetText As String, Optional NewText As String = "")

 Dim FSO         As FileSystemObject 'ファイルシステムオブジェクト
 Dim Txt         As TextStream       'テキストストリームオブジェクト
 Dim buf_strTxt  As String           '読み込みバッファ

 On Error GoTo Func_Err:

 'オブジェクト作成
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(fileName, ForReading)

 '全文読み込み
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '元ファイルをリネームして、テンポラリファイル作成
  Name fileName As fileName & "_"

  '置換処理
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  '書込み用テキストファイル作成
   Set Txt = FSO.CreateTextFile(fileName, True)
  '書込み
  Txt.Write buf_strTxt
  Txt.Close

  'テンポラリファイルを削除
  FSO.DeleteFile fileName & "_"

'終了処理
Func_Exit:
    Set Txt = Nothing
    Set FSO = Nothing
    Exit Function

Func_Err:
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
End Function


'**************************************************************************************************
' * VBAでExcelのコメントを一括で自動サイズにしてカッコよくする
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setComment(Optional BgColorVal, Optional FontVal, Optional FontColorVal = 8421504, Optional FontSizeVal = 9)
  Dim cl As Range
  
  For Each cl In Selection
    DoEvents
    If Not cl.Comment Is Nothing Then
      With cl.Comment.Shape
        'サイズ設定
        .TextFrame.AutoSize = True
        .TextFrame.Characters.Font.Size = FontSizeVal
        '.TextFrame.Characters.Font.Color = FontColorVal

        '形状を角丸四角形に変更
        .AutoShapeType = msoShapeRectangle

        '色
        .line.ForeColor.RGB = RGB(128, 128, 128)
        .Fill.ForeColor.RGB = BgColorVal

        '影 透過率 30%、オフセット量 x:1px,y:1px
        .Shadow.Transparency = 0.3
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 1

        ' 太字解除、中央揃え
        .TextFrame.Characters.Font.Bold = False
        .TextFrame.HorizontalAlignment = xlLeft
        .TextFrame.Characters.Font.Name = FontVal

        ' セルに合わせて移動する
        .Placement = xlMove
      End With
    End If
  Next cl
End Function


'**************************************************************************************************
' * クリップボードクリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetClipboard()
'  OpenClipboard 0
'  EmptyClipboard
'  CloseClipboard
End Function


'**************************************************************************************************
' * 選択セルの行背景解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLineColor(SetArea As String)
  ActiveWorkbook.ActiveSheet.Range(SetArea).Select

  '条件付き書式をクリア
  Selection.FormatConditions.delete
End Function


'**************************************************************************************************
' * リンク解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLink()
  Dim Links As Variant
  Dim i As Integer

  Set wb = ActiveWorkbook
  Links = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks) 'ブックの中にあるリンク

  If IsArray(Links) Then
    For i = 1 To UBound(Links)
      ActiveWorkbook.BreakLink Links(i), xlLinkTypeExcelLinks 'リンク解除
    Next i
  End If
End Function


'**************************************************************************************************
' * スリープ処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  DoEvents
  Application.Wait [Now()] + timeVal / 86400000
  DoEvents
End Function


'**************************************************************************************************
' * 罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 罫線_クリア(Optional SetArea As Range)
  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  End If
End Function

'==================================================================================================
Function 罫線_表(Optional SetArea As Range, Optional LineColor As Long)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = xlThin
      .Borders(xlEdgeRight).Weight = xlThin
      .Borders(xlEdgeTop).Weight = xlThin
      .Borders(xlEdgeBottom).Weight = xlThin
      
      .Borders(xlInsideVertical).Weight = xlThin
      .Borders(xlInsideHorizontal).Weight = xlHairline

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = xlThin
      .Borders(xlEdgeRight).Weight = xlThin
      .Borders(xlEdgeTop).Weight = xlThin
      .Borders(xlEdgeBottom).Weight = xlThin

      .Borders(xlInsideVertical).Weight = xlThin
      .Borders(xlInsideHorizontal).Weight = xlHairline

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)

        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_囲み(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_格子(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_左(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeLeft).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeLeft).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_右(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_左右(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_上(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeTop).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeTop).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_上下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_垂直(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_水平(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_囲み(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_格子(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_左(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeLeft).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeLeft).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_右(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_左右(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_上(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeTop).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeTop).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_上下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_垂直(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideVertical).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideVertical).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_実線_水平(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With

  End If
End Function

'==================================================================================================
Function 罫線_二重線_囲み(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble
      
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_二重線_左(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_二重線_左右(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_二重線_上(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlDouble
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDouble
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_二重線_下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeBottom).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_二重線_上下(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function 罫線_破線_逆L字(Optional SetArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long
  
  Call 罫線_破線_囲み(SetArea, LineColor, WeightVal)
  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    Set SetArea = SetArea.Offset(1, 1).Resize(SetArea.Rows.count - 1, SetArea.Columns.count - 1)
    Call 罫線_破線_水平(SetArea, LineColor, WeightVal)
    Call 罫線_破線_囲み(SetArea, LineColor, WeightVal)
  Else
    SetArea.Offset(1, 1).Resize(SetArea.Rows.count - 1, SetArea.Columns.count - 1).Select
    Call 罫線_破線_水平(SetArea, LineColor, WeightVal)
    Call 罫線_破線_囲み(SetArea, LineColor, WeightVal)
  
  End If
End Function

'==================================================================================================
Function 罫線_中央線削除_横(Optional SetArea As Range)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  End If
End Function

'==================================================================================================
Function 罫線_中央線削除_縦(Optional SetArea As Range)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlInsideVertical).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlNone
    End With
  End If
End Function


'**************************************************************************************************
' * カラム幅設定 / 取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getColumnWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String

  endColLine = Cells(5, Columns.count).End(xlToLeft).Column
  
  For colLine = 1 To endColLine
    Cells(1, colLine) = Cells(1, colLine).ColumnWidth
  Next
End Function

'==================================================================================================
Function setColumnWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String

  endColLine = Cells(1, Columns.count).End(xlToLeft).Column
  
  For colLine = 1 To endColLine
    Cells(1, colLine).ColumnWidth = Cells(1, colLine)
  Next
End Function


'**************************************************************************************************
' * ページのステータス確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getPageStatusCode(ByVal strURL As String) As Integer
  'Dim Http As New WinHttpRequest
  Dim Http As Object
  Dim statusCode As Integer
  
  Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  
  On Error GoTo catchError
  Call Library.showDebugForm("URL：" & strURL)
  If strURL = "" Then
    Exit Function
  End If
  
  With Http
    .Open "HEAD", strURL, False
    
    If BK_setVal("proxyURL") <> "" Then
      .SetProxy 2, BK_setVal("proxyURL") & ":" & BK_setVal("proxyPort")
    End If
    If BK_setVal("proxyUser") <> "" Then
      .setProxyCredentials BK_setVal("proxyUser"), BK_setVal("proxyPasswd")
    End If
    
    .Send
    Call Library.showDebugForm("Status：" & .Status)
    If .Status = 301 Or .Status = 302 Then
      Call Library.showDebugForm("GetAllResponseHeaders：" & .GetAllResponseHeaders)
    Else
      statusCode = .Status
    End If
  End With
  getPageStatusCode = statusCode
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Set Http = Nothing
  Call Library.showDebugForm(Err.Number & "：" & Err.Description)
  
  getPageStatusCode = 404
End Function


