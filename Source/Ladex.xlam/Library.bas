Attribute VB_Name = "Library"
Option Explicit

'**************************************************************************************************
' * 参照設定、定数宣言
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' Windows APIの利用--------------------------------------------------------------------------------
#If VBA7 And Win64 Then
  'ディスプレイの解像度取得用
  Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'Sleep関数の利用
  Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As LongPtr)

  'クリップボード関連
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

  'Shell関数で起動したプログラムの終了を待つ
  Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


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
#End If

Private Const PROCESS_QUERY_INFORMATION = &H400&
Private Const STILL_ACTIVE = &H103&


'ワークブック用変数------------------------------
'ワークシート用変数------------------------------
'グローバル変数----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

Public CalculatFlg As Boolean


'アクティブセルの取得
Dim SelectionCell As String
Dim SelectionSheet As String

' PC、Office等の情報取得用連想配列
Public MachineInfo As Object

'ランダム関数用
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

Public Const JapaneseCharacters = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ"
Public Const JapaneseCharactersCommonUse = "雨学空金青林画岩京国姉知長直店東歩妹明門夜委育泳岸苦具幸始使事実者昔取受所注定波板表服物放味命油和英果芽官季泣協径固刷参治周松卒底的典毒念府法牧例易往価河居券効妻枝舎述承招性制版肥非武沿延拡供呼刻若宗垂担宙忠届乳拝並宝枚依押奇祈拠況屈肩刺沼征姓拓抵到突杯泊拍迫彼怖抱肪茂炎欧殴"
Public Const MachineDependentCharacters = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㍉纊褜鍈銈蓜俉炻昱棈鋹曻彅丨仡仼伀伃伹佖侒侊侚侔俍偀倢俿倞偆偰偂傔"

Public ThisBook As Workbook


'スタイル関連------------------------------------
'Public useStyle()           As Variant





'**************************************************************************************************
' * エラー時の処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function errorHandle()
  On Error Resume Next
  
  Call Library.endScript
  Call Ctl_ProgressBar.showEnd
  Call Library.showDebugForm(funcName, , "end1")
  Call init.unsetting
  
End Function

'**************************************************************************************************
' * 画面描写制御開始
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()
  Const funcName As String = "Library.startScript"

  On Error Resume Next
  Call Library.showDebugForm(funcName, , "function")
  
  'アクティブセルの取得
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
    SelectionSheet = ActiveWorkbook.ActiveSheet.Name
  End If

  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  If Application.Calculation = xlCalculationManual Then
    CalculatFlg = False
  Else
    Application.Calculation = xlCalculationManual
    CalculatFlg = True
  End If
  Call Library.showDebugForm("CalculatFlg", CalculatFlg, "debug")

  Application.ScreenUpdating = False              'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.EnableEvents = False                'マクロ動作自体で別のイベントが生成されるのを抑制する
  
  Application.DisplayAlerts = False               '確認メッセージを出さない
  'Application.StatusBar = "処理中・・・"         'ステータスバーに処理中を表示

'  If runFlg = True Then
'    Application.Interactive = False                 'マクロ動作中に一切のキーやマウス操作を制限する
'    Application.Cursor = xlWait                     'マクロ動作中はマウスカーソルを「砂時計」にする
'  End If

End Function

'**************************************************************************************************
' * 画面描写制御終了
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional reCalflg As Boolean = False, Optional flg As Boolean = False)
  Const funcName As String = "Library.endScript"

  On Error Resume Next
  Call Library.showDebugForm(funcName, , "function")

  '強制的に再計算させる
  If reCalflg = True Then
    Application.CalculateFull
  Else
    ActiveSheet.Calculate
  End If
  
  Call Library.showDebugForm("CalculatFlg", CalculatFlg, "debug")
  If CalculatFlg = True Then
    Application.Calculation = xlCalculationAutomatic  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  End If

 'アクティブセルの選択
  If SelectionCell <> "" And flg = True Then
    ActiveWorkbook.Worksheets(SelectionSheet).Select
    ActiveWorkbook.Range(SelectionCell).Select
  End If

  Application.ScreenUpdating = True                 'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.EnableEvents = True                   'マクロ動作自体で別のイベントが生成されるのを抑制する
  
  Application.Interactive = True                    'マクロ動作中に一切のキーやマウス操作を制限する
  Application.Cursor = xlDefault                    'マクロ動作終了後はマウスカーソルを「デフォルト」にもどす
  Application.StatusBar = False                     'マクロ動作終了後はステータスバーを「デフォルト」にもどす
  Application.DisplayAlerts = True                  '確認メッセージを出さない
End Function

'**************************************************************************************************
' * シートの存在確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkSheetExists(sheetName) As Boolean
  Dim tempSheet As Object
  Dim Result As Boolean

  Result = False
  For Each tempSheet In Sheets
    If LCase(sheetName) = LCase(tempSheet.Name) Then
      Result = True
      Exit For
    End If
  Next
  chkSheetExists = Result
End Function

'**************************************************************************************************
' * 処理完了まで待機
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShellEnd(ProcessID As Long)
  Dim hProcess As Long, EndCode As Long, EndRet As Long

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
Function chkShapeName(ShapeName As String, Optional targetSheet As Worksheet) As Boolean
  Dim objShp As Shape
  Dim Result As Boolean

  Result = False
  
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  
  For Each objShp In targetSheet.Shapes
    If objShp.Name = ShapeName Then
      Result = True
      Exit For
    End If
  Next
  chkShapeName = Result
End Function


'**************************************************************************************************
' * 除外シート判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(chkSheetName As String) As Boolean
 Dim Result As Boolean
 Dim sheetName As Variant

  For Each sheetName In Range("ExcludeSheet")
    If sheetName = chkSheetName Then
      Result = True
      Exit For
    Else
      Result = False
    End If
  Next
  chkExcludeSheet = Result
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
' * 配列に値が存在するかどうか
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkArrayVal(arrayTmp As Variant, chkVal As String) As Boolean
  Dim filterVal As Variant

  On Error GoTo catchError
  filterVal = Filter(arrayTmp, chkVal, True)
  If (UBound(filterVal) <> -1) Then
    chkArrayVal = True
  Else
    chkArrayVal = False
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
  chkArrayVal = True
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
  Const funcName As String = "Library.chkHeader"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
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
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
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

'==================================================================================================
Function chkIsOpen(targetBookName As String, Optional fileCnt As Integer = 0) As Boolean
  Dim openWorkbook As Workbook
  Dim chkFlg As Boolean
  
  Const funcName As String = "Library.chkIsOpen"

  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("targetBookName", targetBookName, "debug")
  
  chkFlg = False
  fileCnt = 0
  For Each openWorkbook In Workbooks
    Call Library.showDebugForm("openWorkbook  ", openWorkbook.Name, "debug")
    
    If InStr(targetBookName, "*") > 0 Then
      If openWorkbook.Name Like targetBookName Then
        fileCnt = fileCnt + 1
      End If
    Else
      If openWorkbook.Name = targetBookName Then
        chkFlg = True
        fileCnt = 1
        Exit For
      End If
    End If
  Next
  
  
  Call Library.showDebugForm("fileCnt", fileCnt, "debug")
  Call Library.showDebugForm("isOpen ", chkFlg, "debug")
  
  Call Library.showDebugForm(funcName, , "end1")
  
  chkIsOpen = chkFlg
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

  If dicVal("debugMode") = "develop" Then
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
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkFileExists(targetPath As String)
  Dim FSO As Object
  Dim retVal As Boolean

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FileExists(targetPath) Then
    retVal = True
  Else
    retVal = False
  End If
  Set FSO = Nothing
  chkFileExists = retVal
End Function

'**************************************************************************************************
' * ディレクトリの存在確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
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
' * 値の型チェック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkTypeName(targetVal As Variant, permitType As String, Optional regPattern As String)
  Dim regexp
  Dim resultFlg As Boolean

  Const funcName As String = "Library.chkTypeName"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Set regexp = CreateObject("VBScript.RegExp")
  resultFlg = False

  If targetVal = "" Then
    resultFlg = True
    GoTo Lbl_endSelect
  End If
  regexp.Global = True              '文字列全体を検索

  Select Case permitType
    Case "int"        '数字
      regexp.Pattern = "^[0-9]+$"
      resultFlg = regexp.test(targetVal)

    Case "string"     '英小文字
      regexp.IgnoreCase = False
      regexp.Pattern = "^[a-z]+$"
      resultFlg = regexp.test(targetVal)

    Case "STRING"     '英大文字
      regexp.IgnoreCase = False
      regexp.Pattern = "^[A-Z]+$"
      resultFlg = regexp.test(targetVal)

    Case "String"     '英字(大小区別なし)
      regexp.IgnoreCase = True
      regexp.Pattern = "^[a-zA-Z]+$"
      resultFlg = regexp.test(targetVal)

    Case "reg"        '正規表現
      regexp.IgnoreCase = True
      regexp.Pattern = regPattern
      resultFlg = regexp.test(targetVal)

    Case "date"
      resultFlg = IsDate(targetVal)
  End Select
  Set regexp = Nothing

Lbl_endSelect:
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("targetval ", targetVal, "info")
  Call Library.showDebugForm("regPattern", regPattern, "info")
  Call Library.showDebugForm("resultFlg ", resultFlg, "info")

  If resultFlg = True Then
    chkTypeName = False
  Else
    chkTypeName = True
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * ByteからKB,MB,GBへ変換
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convscale(ByVal lngVal As Long) As String
  Dim convVal As String

  If lngVal >= 1024 ^ 3 Then
    convVal = Round(lngVal / (1024 ^ 3), 2) & " GB"
  ElseIf lngVal >= 1024 ^ 2 Then
    convVal = Round(lngVal / (1024 ^ 2), 2) & " MB"
  ElseIf lngVal >= 1024 Then
    convVal = Round(lngVal / (1024), 2) & " KB"
  Else
    convVal = lngVal & " B"
  End If
  convscale = convVal
End Function

'**************************************************************************************************
' * 固定長文字列に変換
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convFixedLength(targetVal As String, lengs As Long, addString As String, Optional addType As Boolean = True) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  Do While LenB(StrConv(targetVal, vbFromUnicode)) <= lengs
    If addType = True Then
      targetVal = targetVal & addString
    Else
      targetVal = addString & targetVal
    End If
  Loop
  convFixedLength = targetVal
End Function


'**************************************************************************************************
' * キャメルケースをスネークケースに変換
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function covCamelToSnake(ByVal val As String, Optional ByVal isUpper As Boolean = False) As String
  Dim ret As String
  Dim i As Long, Length As Long

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

  val = LCase(val)
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
' * 半角を全角に変換する(英数字、カタカナ)
' *
' * @link   http://officetanaka.net/excel/function/tips/tips45.htm
'**************************************************************************************************
Function convHan2Zen(Text As String) As String
  Dim i As Long, buf As String
  Dim c As Range
  Dim rData As Variant, ansData As Variant
  Const funcName As String = "Library.convHan2Zen"
  
  convHan2Zen = StrConv(Text, vbWide)
End Function

'**************************************************************************************************
' * 全角を半角に変換する(英数字、カタカナ)
' *
' * @link   http://officetanaka.net/excel/function/tips/tips45.htm
'**************************************************************************************************
Function convZen2Han(ByVal Text As String) As String
  Dim i As Long, buf As String
  Dim c As Range
  Dim covText As String
  Const funcName As String = "Library.convZen2Han"
  
  For i = 1 To Len(Text)
    buf = Mid(Text, i, 1)
    If buf Like "[Ａ-ｚ０-９]" Or _
      buf Like "[－＝！．｛｝（）／]" Then
      covText = covText & StrConv(buf, vbNarrow)
        
    ElseIf buf Like "[ｱ-ﾝ]" Then
      covText = covText & StrConv(buf, vbWide)
    
    ElseIf buf = "," Then
      covText = covText & "，"
    
    Else
      covText = covText & buf
    End If
    DoEvents
  Next i
  
  Call Library.showDebugForm(funcName, covText, "debug")
  convZen2Han = covText
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
Function convBase64EncodeForFile(ByVal FilePath As String) As String
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
    .LoadFromFile FilePath
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
' * @link   https://www.ka-net.org/office/of32.html
'**************************************************************************************************
Function convURLEncode(ByVal str As String) As String
  Dim EncodeURL As String

#If VBA7 And Win64 Then
  Dim d As Object
  Dim elm As Object
  
  str = Replace(str, "\", "\\")
  str = Replace(str, "'", "\'")
  Set d = CreateObject("htmlfile")
  Set elm = d.createElement("span")
  elm.setAttribute "id", "result"
  d.body.appendChild elm
  d.parentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & str & "');", "JScript"
  EncodeURL = elm.innerText
#Else
  With CreateObject("ScriptControl")
    .Language = "JScript"
    EncodeURL = .CodeObject.encodeURIComponent(str)
  End With
#End If

  convURLEncode = EncodeURL
End Function


'**************************************************************************************************
' * URLデコード
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLDecode(ByVal str As String) As String
  Dim DecodeURL As String

#If VBA7 And Win64 Then
  Dim d As Object
  Dim elm As Object
  
  str = Replace(str, "\", "\\")
  str = Replace(str, "'", "\'")
  Set d = CreateObject("htmlfile")
  Set elm = d.createElement("span")
  elm.setAttribute "id", "result"
  d.body.appendChild elm
  d.parentWindow.execScript "document.getElementById('result').innerText = decodeURIComponent('" & str & "');", "JScript"
  DecodeURL = elm.innerText
#Else
  With CreateObject("ScriptControl")
    .Language = "JScript"
    DecodeURL = .CodeObject.decodeURIComponent(str)
  End With
#End If
  
  convURLDecode = DecodeURL
End Function


'**************************************************************************************************
' * Unicodeエスケープ
' *
' * @link   https://qiita.com/mima_ita/items/8fc5fab7259835e4bcdd
'**************************************************************************************************
Public Function convUnicodeEscape(ByVal StringToEncode As String) As String
  Dim i As Integer
  Dim acode As Integer
  Dim char As String, escape As String
  
  Const funcName As String = "Library.convUnicodeEscape"
  
  Call Library.showDebugForm(funcName, , "start1")
  escape = StringToEncode

  For i = Len(escape) To 1 Step -1
    acode = AscW(Mid$(escape, i, 1))
    Call Library.showDebugForm("対象文字列", Mid$(escape, i, 1) & "<:>" & acode, "debug")
    
    Select Case acode
      Case 48 To 57, 65 To 90, 97 To 122, 123, 125
      
  
      Case 32
        escape = Left$(escape, i - 1) & "%20" & Mid$(escape, i + 1)
  
      Case Else
        char = Hex$(acode)
        If Len(char) > 2 Then
          If Len(char) = 3 Then
            char = "0" & char
          End If
          escape = Left$(escape, i - 1) & "\u" & char & Mid$(escape, i + 1)
        Else
          If Len(char) = 1 Then
            char = "0" & char
          End If
          escape = Left$(escape, i - 1) & "\" & char & Mid$(escape, i + 1)
        End If
        
        Call Library.showDebugForm("escape", escape, "debug")
    End Select
  Next
  
  convUnicodeEscape = LCase(escape)
  
  Call Library.showDebugForm(funcName, , "end")
End Function


'**************************************************************************************************
' * Unicodeエスケープ
' *
' * @link   http://tech7.blog.shinobi.jp/vba/unicodeエスケープされた文字列を文字に戻す方法
'**************************************************************************************************
Public Function convUnicodeunEscape(ByVal strTarget As String) As String
  Dim str As String
  Dim strRet As String
  Dim lngPos As Long
  Dim lngStart As Long
  Dim strTmp As String
 
 
  str = strTarget
  lngPos = 0
  Do
    lngStart = lngPos
    lngPos = InStr(1, str, "\u")
    
    If lngPos > 0 Then
     strRet = strRet & Mid(str, 1, lngPos - 1)
     strTmp = Mid(str, lngPos, 6)
    
     strTmp = Replace(strTmp, "\u", "&H")
     strRet = strRet & ChrW(strTmp)
    
     str = Mid(str, lngPos + 6)
    
    Else
     strRet = strRet & str
    
     Exit Do
    End If
  Loop
  convUnicodeunEscape = strRet
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
Function delSheetData(Optional targetSheet As Worksheet, Optional line As Long, Optional delImgFlg As Boolean = False)
  Dim shp As Shape
  Const funcName As String = "Library.delSheetData"

  Call Library.showDebugForm(funcName, , "start1")
  
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  Call Library.showDebugForm("sheetName", targetSheet.Name, "debug")
  Call Library.showDebugForm("delLine  ", line, "debug")

  If targetSheet.FilterMode = True Or targetSheet.AutoFilterMode = True Then
    targetSheet.AutoFilterMode = False
  End If

  If line <> 0 Then
    targetSheet.Rows(line & ":" & Rows.count).delete Shift:=xlUp
    targetSheet.Rows(line & ":" & Rows.count).NumberFormatLocal = "G/標準"
    targetSheet.Rows(line & ":" & Rows.count).style = "Normal"
  Else
    targetSheet.Cells.delete Shift:=xlUp
    targetSheet.Cells.NumberFormatLocal = "G/標準"
    targetSheet.Cells.style = "Normal"
  End If
  DoEvents

  If delImgFlg = True Then
    For Each shp In ActiveSheet.Shapes
    shp.Select
      If shp.Type = 11 Then shp.delete
    Next shp
  End If
  
  Call Library.showDebugForm(funcName, , "end1")
  
'  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 1, "データ消去：" & targetSheet.name)
  
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
Function セルの名称設定削除()
  Dim Name As Object

  Const funcName As String = "Library.セルの名称設定削除"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  On Error Resume Next
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.delete
      Call Library.showDebugForm("Name", Name.Name, "debug")
    End If
  Next
  
  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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
  Const funcName As String = "Library.execCopy"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call Library.showDebugForm("  コピー元：" & srcPath)
  Call Library.showDebugForm("  コピー先：" & dstPath)

  If chkFileExists(srcPath) = False Then
    Call Library.showNotice(404, "コピー元", True)
  End If

  If chkDirExists(getParentDir(dstPath)) = False Then
    Call Library.execMkdir(getParentDir(dstPath))
  End If
  FSO.CopyFile srcPath, dstPath
  Set FSO = Nothing
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * ファイル移動
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execMove(srcPath As String, dstPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execMove"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call Library.showDebugForm("  移動元：" & srcPath)
  Call Library.showDebugForm("  移動先：" & dstPath)

  If chkFileExists(srcPath) = False Then
    Call Library.showNotice(404, "移動元", True)
  End If

  FSO.MoveFile srcPath, dstPath
  Set FSO = Nothing
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * ディレクトリ削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execDeldir(srcPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execDeldir"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call Library.showDebugForm("  削除対象：" & srcPath)

  If srcPath Like "*[*]*" Then
  ElseIf chkDirExists(srcPath) = False Then
    Call Library.showNotice(404, "削除対象", True)
  End If

  FSO.DeleteFolder srcPath
  Set FSO = Nothing
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * ファイル削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execDel(srcPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execDel"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call Library.showDebugForm("削除対象", srcPath, "debug")

  If srcPath Like "*[*]*" Then
  ElseIf chkFileExists(srcPath) = False Then
    Call Library.showNotice(404, "削除対象", True)
  End If

  FSO.DeleteFile srcPath
  Set FSO = Nothing
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
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

  errFlg = False
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("変更元", srcPath)
  Call Library.showDebugForm("旧名称", oldFileName)
  Call Library.showDebugForm("新名称", fileName)

  If chkFileExists(srcPath & "\" & oldFileName) = False Then
    If IsMissing(errMeg) Then
      Call Library.showNotice(404, "変更元", True)
    Else
      errMeg = "変更元のファイルがありません[" & oldFileName & "]"
      errFlg = True
    End If

  End If
  If chkFileExists(srcPath & "\" & fileName) = True Then
    If IsMissing(errMeg) Then
      Call Library.showNotice(414, fileName, True)
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
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  errMeg = Err.Description
  execRename = False
End Function

'**************************************************************************************************
' * MkDirで階層の深いフォルダーを作る
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execMkdir(fullPath As String)

  If chkDirExists(fullPath) Then
    Exit Function
  End If

  Call Library.showDebugForm("execMkdir", fullPath, "debug")
  Call chkParentDir(fullPath)
End Function

'==================================================================================================
Private Function chkParentDir(TargetFolder)
  Dim ParentFolder As String, FSO As Object

  Const funcName As String = "Library.chkParentDir"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("TargetFolder", TargetFolder, "debug")

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
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * zip圧縮/解凍
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execCompress(srcPath As String, zipFilePath As String) As Boolean
  'Dim sh  As New IWshRuntimeLibrary.WshShell
  Dim Sh
  Dim ex  As WshExec
  Dim cmd As String
  Set Sh = CreateObject("WScript.Shell")
  Call Library.showDebugForm("対象ディレクトリ：" & srcPath)
  Call Library.showDebugForm("zipファイル     ：" & zipFilePath)

  If chkDirExists(srcPath) = False Then
    Call Library.showNotice(403, "対象ディレクトリ", True)
  End If

  '半角スペースをバッククォートでエスケープ
  srcPath = Replace(srcPath, " ", "` ")
  zipFilePath = Replace(zipFilePath, " ", "` ")

  cmd = "Compress-Archive -Path " & srcPath & " -DestinationPath " & zipFilePath & " -Force"
  Call Library.showDebugForm("cmd：" & cmd)
  Set ex = Sh.exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & cmd)

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
  'Dim sh As New IWshRuntimeLibrary.WshShell
  Dim Sh
  Dim ex As WshExec
  Dim cmd As String

  Set Sh = CreateObject("WScript.Shell")
  Call Library.showDebugForm("zipファイル     ", zipFilePath)
  Call Library.showDebugForm("対象ディレクトリ", dstPath)

  If chkFileExists(zipFilePath) = False Then
    Call Library.showNotice(404, "解凍対象", True)
  End If
  If chkDirExists(dstPath) = False Then
    Call Library.showNotice(403, "解凍先", True)
  End If

  '半角スペースをバッククォートでエスケープ
  zipFilePath = Replace(zipFilePath, " ", "` ")
  dstPath = Replace(dstPath, " ", "` ")

  cmd = "Expand-Archive -Path " & zipFilePath & " -DestinationPath " & dstPath & " -Force"
  Call Library.showDebugForm("cmd：" & cmd)
  Set ex = Sh.exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & cmd)

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

  'OSのバージョン取得----------------------------
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

  'Excelのバージョン取得-------------------------
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

  'PCの情報--------------------------------------
  MachineInfo.add "UserName", WshNetworkObject.userName
  MachineInfo.add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.add "UserDomain", WshNetworkObject.UserDomain

  '画面の解像度等取得----------------------------
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
Function getLength(targetVal As String, Optional charType As String = "半角")
  Dim inputLen As Long
  Const funcName As String = "Library.getLength"
  
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm("targetVal", targetVal, "debug")
'  Call Library.showDebugForm("charType ", charType, "debug")
'  Call Library.showDebugForm("文字数   [Len]", Len(targetVal), "debug")
'  Call Library.showDebugForm("バイト数[LenB]", LenB(StrConv(targetVal, vbFromUnicode)), "debug")
  
  Select Case charType
    Case "半角", "全角"
      inputLen = LenB(StrConv(targetVal, vbFromUnicode))
    Case "文字数"
      inputLen = Len(targetVal)
  End Select
  
'  Call Library.showDebugForm("inputLen", inputLen, "debug")
  getLength = inputLen
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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
  Call Library.showDebugForm("R1C1Top     ：" & R1C1Top)
  Call Library.showDebugForm("R1C1Left    ：" & R1C1Left)
  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("Rng.Address ：" & Rng.Address)
  Call Library.showDebugForm("ActvCellTop ：" & ActvCellTop)
  Call Library.showDebugForm("ActvCellLeft：" & ActvCellLeft)
End Function

'**************************************************************************************************
' * セルの選択範囲取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getCellSelectArea(Optional startLine As Long, Optional endLine As Long, Optional startColLine As Long, Optional endColLine As Long)
  Dim tmpLine As Long
  
  Const funcName As String = "Library.getCellSelectArea"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("Selection   ", Selection.CountLarge, "debug")
  '----------------------------------------------

  '選択範囲がある場合----------------------------
  If Selection.CountLarge > 1 Then
    startLine = Selection(1).Row
    endLine = Selection(Selection.count).Row
    tmpLine = Range("A1").SpecialCells(xlLastCell).Row
    If endLine > tmpLine Then
      endLine = tmpLine
    End If
    
    
    startColLine = Selection.Column
    endColLine = Selection.Column + Selection.Columns.count - 1
    tmpLine = Range("A1").SpecialCells(xlLastCell).Column
    If endColLine > tmpLine Then
      endColLine = tmpLine
    End If
    
    
  
  '選択範囲がない場合----------------------------
  Else
    startLine = 1
    endLine = Range("A1").SpecialCells(xlLastCell).Row
    
    startColLine = 1
    endColLine = Range("A1").SpecialCells(xlLastCell).Column
  End If
  
  If endLine = 0 Then
    endLine = startLine
  End If
  
  Call Library.showDebugForm("startLine   ", startLine, "debug")
  Call Library.showDebugForm("endLine     ", endLine, "debug")
  Call Library.showDebugForm("startColLine", startColLine, "debug")
  Call Library.showDebugForm("endColLine  ", endColLine, "debug")

  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * 列名から列番号を求める
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long
  getColumnNo = Range(targetCell & ":" & targetCell).Column
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
'  If setColorValue = False Then
'    setColorValue = colorValue
'  End If
  getColor = setColorValue
End Function

'**************************************************************************************************
' * フォントダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFont(FontName As String, FontSize As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Application.Dialogs(xlDialogActiveCellFont).Show FontName, "レギュラー", FontSize
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
Function getDirPath(CurrentDirectory As String, Optional title As String, Optional setRegPathName As String = "")
  Dim tmpPath As String

  If setRegPathName <> "" Then
    tmpPath = Library.getRegistry("targetInfo", setRegPathName)
    If tmpPath <> "" Then
      CurrentDirectory = tmpPath
    End If
  End If
  
  With Application.FileDialog(msoFileDialogFolderPicker)
    If Library.chkDirExists(CurrentDirectory) = True Then
      .InitialFileName = CurrentDirectory & "\"
    Else
      .InitialFileName = ActiveWorkbook.path
    End If

    .AllowMultiSelect = False
    If title <> "" Then
      .title = title & "の場所を選択してください"
    Else
      .title = "フォルダーを選択してください"
    End If
    
    If .Show = True Then
      Call Library.setRegistry("targetInfo", setRegPathName, .SelectedItems(1))
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
  Dim FilePath As String
  Dim Result As Long
  Dim fileName As Variant

  fileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm,すべてのファイル, *.*", _
      FilterIndex:=FileTypeNo)

  If fileName <> "False" Then
    getSaveFilePath = FilePath
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
  Dim FilePath As String
  Dim Result As Long

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
      .InitialFileName = ActiveWorkbook.path & "\" & fileName
    End If

    '表示形式の設定
    .InitialView = msoFileDialogViewWebView

    'ダイアログ ボックスのタイトル設定
    .title = title & "を選択してください"

    If .Show = -1 Then
      FilePath = .SelectedItems(1)
    Else
      FilePath = ""
    End If
  End With
  getFilePath = FilePath
End Function

'**************************************************************************************************
' * 複数ファイル選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, title As String, fileType As String, Optional setRegPathName As String = "")
  Dim FilePath() As Variant
  Dim tmpPath As String
  Dim Result As Long, i As Integer

  If setRegPathName <> "" Then
    tmpPath = Library.getRegistry("targetInfo", setRegPathName)
    If tmpPath <> "" Then
      CurrentDirectory = tmpPath
    End If
  End If
  
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

      Case "mov"
        .Filters.add "動画ファイル", "*.mp4"

      Case Else
        .Filters.add "すべてのファイル", "*.*"
    End Select
    '.FilterIndex = FileTypeNo

    '表示するフォルダ
    .InitialFileName = CurrentDirectory & "\"

    '表示形式の設定
    .InitialView = msoFileDialogViewWebView

    'ダイアログ ボックスのタイトル設定
    .title = title

    If .Show = -1 Then
      Call Library.setRegistry("targetInfo", setRegPathName, Library.getFileInfo(.SelectedItems(1), , "CurrentDir"))
      
      ReDim Preserve FilePath(.SelectedItems.count - 1)
      For i = 1 To .SelectedItems.count
        FilePath(i - 1) = .SelectedItems(i)
      Next i
    Else
      ReDim Preserve FilePath(0)
      FilePath(0) = ""
    End If
  End With
  getFilesPath = FilePath
End Function

'**************************************************************************************************
' * ディレクトリ内のファイル一覧取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileList(path As String, fileName As String)
  Dim f As Object, cnt As Long
  Dim list() As String

  cnt = 0
  Call Library.showDebugForm("Path", path, "info")
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(path).Files
      If f.Name Like fileName Then
        ReDim Preserve list(cnt)
        list(cnt) = f.Name
        cnt = cnt + 1
      End If
    Next f
  End With
  getFileList = list
End Function

'==================================================================================================
Function getFilePath2LikeFileName(path As String, fileName As String, Optional perfectMatchFlg As Boolean = False)
  Dim f As Object
  Dim retVal As String
  Const funcName As String = "Library.getFilePath2likeFileName"

  Call Library.showDebugForm("Path", path, "info")
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(path).Files
      If f.Name Like fileName And perfectMatchFlg = False Then
        retVal = path & "\" & f.Name
        Exit For
      ElseIf f.Name = fileName And perfectMatchFlg = True Then
        retVal = path & "\" & f.Name
        Exit For
      End If
    Next f
  End With
  getFilePath2LikeFileName = retVal
End Function

'**************************************************************************************************
' * ファイル情報取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileInfo(targetFilePath As String, Optional FileInfo As Object, Optional getType As String)
  Dim FSO As Object
  Dim fileObject As Object
  Dim sp As Shape

  Call Library.showDebugForm("targetFilePath", targetFilePath, "debug")
  If Library.chkFileExists(targetFilePath) = False Then
    Exit Function
  End If
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set FileInfo = Nothing
  Set FileInfo = CreateObject("Scripting.Dictionary")
  

  '作成日時
  FileInfo.add "createAt", Format(FSO.GetFile(targetFilePath).DateCreated, "yyyy/mm/dd hh:nn:ss")

  '更新日時
  FileInfo.add "updateAt", Format(FSO.GetFile(targetFilePath).DateLastModified, "yyyy/mm/dd hh:nn:ss")

  'ファイルサイズ
  FileInfo.add "size", FSO.GetFile(targetFilePath).Size

  'ファイルの種類
  FileInfo.add "type", FSO.GetFile(targetFilePath).Type

  '拡張子
  FileInfo.add "extension", FSO.GetExtensionName(targetFilePath)

  'ファイル名
  FileInfo.add "fileName", FSO.GetFile(targetFilePath).Name

  'ファイルが存在するフォルダ
  FileInfo.add "CurrentDir", FSO.GetFile(targetFilePath).ParentFolder

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

          FileInfo.add "width", CLng(.Width * 4 / 3)
          FileInfo.add "height", CLng(.Height * 4 / 3)
          .delete
      End With

    Case "bmp", "jpg", "jpeg", "gif", "emf", "ico", "rle", "wmf"
      Set fileObject = LoadPicture(targetFilePath)
      FileInfo.add "width", fileObject.Width
      FileInfo.add "height", fileObject.Height
      Set fileObject = Nothing

    Case Else
  End Select

  Set FSO = Nothing
  If getType <> "" Then
    getFileInfo = FileInfo(getType)
    Set FileInfo = Nothing
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
' * シートリスト取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSheetList(columnName As String)
  Dim i As Long
  Dim sheetName As Object
  Const funcName As String = "Library.getSheetList"

  i = 3
  If columnName = "" Then
    columnName = "E"
  End If

  Call Library.showDebugForm(funcName, , "start1")

  '現設定値のクリア
  Worksheets("設定").Range(columnName & "3:" & columnName & "100").Select
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle = xlNone
  Selection.Borders(xlEdgeTop).LineStyle = xlNone
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  Selection.Borders(xlInsideVertical).LineStyle = xlNone
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With

  For Each sheetName In ActiveWorkbook.Sheets
    'シート名の設定
    Worksheets("設定").Range(columnName & i).Select
    Worksheets("設定").Range(columnName & i) = sheetName.Name

    ' セルの背景色解除
    With Worksheets("設定").Range(columnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' シート色と同じ色をセルに設定
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("設定").Range(columnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    '罫線の設定
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    i = i + 1
  Next

  Worksheets("設定").Range(columnName & "3").Select
  Call endScript
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
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
' * デバッグ用表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(ByVal meg1 As String, Optional meg2 As Variant, Optional LogLevel As String)
  Dim runTime As String
  Dim StartUpPosition As Long
  Const funcName As String = "Library.showDebugForm"

  On Error GoTo catchError
  runTime = Format(Now(), "yyyy-mm-dd hh:nn:ss")

  Select Case LogLevel
    
    Case "Error"
      LogLevel = 1
      meg1 = "  Error   ：" & Replace(meg1, vbNewLine, " ")

    Case "warning"
      LogLevel = 2
      meg1 = "  Warning ：" & Replace(meg1, vbNewLine, " ")
    
    Case "info"
      LogLevel = 4
      meg1 = "  Info    ：" & Replace(meg1, vbNewLine, " ")

    Case "debug"
      LogLevel = 5
      meg1 = "  Debug   ：" & Replace(meg1, vbNewLine, " ")
    
    Case "start"
      LogLevel = 0
      meg1 = Library.convFixedLength(meg1, 62, "=")
    
    Case "end"
      LogLevel = 0
      meg1 = Library.convFixedLength("", 62, "=")
      
    Case "start1"
      LogLevel = 0
      meg1 = Library.convFixedLength("  " & meg1 & " ", 62, "-")
    
    Case "end1"
      LogLevel = 0
      meg1 = Library.convFixedLength("  ", 62, "-")
      
    Case "function", "function1"
      LogLevel = 0
      meg1 = "  [Function] " & meg1
      
    Case Else
      LogLevel = 6
      meg1 = "  [XXXXXXXX] " & Replace(meg1, vbNewLine, " ")
  End Select

  If IsMissing(meg2) = False Then
    meg1 = meg1 & " : " & Application.WorksheetFunction.Trim(CStr(meg2))
  End If

'  If CLng(LogLevel) <= G_LogLevel Then
'    Call outputLog(runTime, meg1)
'  End If

  If dicVal("debugMode") = "develop" Then
    Debug.Print runTime & "  " & meg1
  End If
  DoEvents

  If LogLevel = 6 Then
    Stop
  End If
  Exit Function

'エラー発生時------------------------------------
catchError:
  Debug.Print runTime & "  " & meg1
  Exit Function
End Function

'**************************************************************************************************
' * ステータスバーにメッセージを表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showError(message As String)
  Dim i As Integer
  Const funcName As String = "Library.showError"

  On Error GoTo catchError

  For i = 0 To 3
    Application.StatusBar = message
    Call Library.waitTime(300)
    
    Application.StatusBar = " "
    Call Library.waitTime(300)
  Next
  
  Application.StatusBar = False

  Exit Function

'エラー発生時------------------------------------
catchError:
  Debug.Print "  [ERROR] " & Err.Description; "  " & message
  Exit Function
End Function

'**************************************************************************************************
' * 処理情報通知
' *
' * Worksheets("info").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional process As String, Optional runEndflg As Boolean)
  Dim message As String, speakerMeg As String, megTitle As String, errLevel As String
  Dim runTime As Date
  Dim endLine As Long

  On Error GoTo catchError
  runTime = Format(Now(), "yyyy-mm-dd hh:nn:ss")

  errLevel = "warning"
  endLine = LadexSh_Notice.Cells(Rows.count, 1).End(xlUp).Row
  message = Application.WorksheetFunction.VLookup(Code, LadexSh_Notice.Range("A2:C" & endLine), 3, False)
  megTitle = Application.WorksheetFunction.VLookup(Code, LadexSh_Notice.Range("A2:C" & endLine), 2, False)
  If megTitle = "" Then megTitle = thisAppName

  message = Replace(message, "%%", process)
  If process = "" Then
    message = Replace(message, "<>", process)
  End If
  If runEndflg = True Then
    speakerMeg = message & vbNewLine & "。処理を中止します"
    errLevel = "Error"
  Else
    speakerMeg = message
  End If

  If message <> "" Then
    message = Replace(message, "<BR>", vbNewLine)
  End If

  If dicVal("debugMode") = "speak" Or dicVal("debugMode") = "develop" Or dicVal("debugMode") = "all" Then
'    Application.Speech.Speak Text:=speakerMeg, SpeakAsync:=True, SpeakXML:=True
  End If

  message = Replace(message, "<", "[")
  message = Replace(message, ">", "]")

  Select Case Code
    Case 0 To 399
      Call MsgBox(message, vbInformation, megTitle)
      errLevel = "end"

    Case 400 To 499
      Call MsgBox(message, vbCritical, megTitle)

    Case 500 To 599
      Call MsgBox(message, vbExclamation, megTitle)

    Case 999

    Case Else
      Call MsgBox(message, vbCritical, megTitle)
  End Select

  message = " [" & Code & "]" & message
  Call Library.showDebugForm(message, , errLevel)

  '画面描写制御終了処理
  If runEndflg = True Then
    Call Library.endScript
    Call Ctl_ProgressBar.showEnd
    Call init.unsetting
    End
  End If

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(message, , errLevel)
  Call MsgBox(message, vbCritical, thisAppName)
End Function

'**************************************************************************************************
' * ランダム
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setStringCnt As Integer, Optional setString As String) As String
  Dim i, n
  Dim str1 As String

  If setString = "" Then
    setString = HalfWidthDigit & HalfWidthCharacters
  End If
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
  Dim makeVal As String, tmpVal As String
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
Function outputLog(runTime As String, message As String)
  Dim fileTimestamp As Date

'  On Error GoTo catchError
  If logFile = "" Then
    Debug.Print runTime & "  " & "ログファイルが設定されていません"
    Exit Function
  
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
    .WriteText runTime & vbTab & message, 1
    .SaveToFile logFile, 2
    .Close
  End With
  
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Debug.Print "[" & Err.Number & "] ログ出力失敗：" & Err.Description
  Debug.Print "[" & Err.Number & "] " & logFile
  Debug.Print "[" & Err.Number & "] " & runTime & vbTab & message
End Function

'==================================================================================================
Function outputText(message As String, outputFilePath As String, Optional encode As String = "sjis")

  With CreateObject("ADODB.Stream")
    If encode = "sjis" Then
      .Charset = "shift_jis"
    ElseIf encode = "utf-8" Then
      .Charset = "UTF-8"
    End If
    
    .Open
    If Library.chkFileExists(outputFilePath) Then
      .LoadFromFile outputFilePath
      .Position = .Size
    End If
    .WriteText message, 1
    .SaveToFile outputFilePath, 2
    .Close
  End With
  
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Debug.Print "[" & Err.Number & "] ファイル出力失敗：" & Err.Description
  Debug.Print "[" & Err.Number & "] " & outputFilePath
  Debug.Print "[" & Err.Number & "] " & message
End Function




'**************************************************************************************************
' * CSV形式ファイルインポート[csv/txt]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @link   https://www.tipsfound.com/vba/18014
'**************************************************************************************************
'==================================================================================================
Function importCsv(FilePath As String, Optional encode As String = "sjis", Optional readLine As Long, Optional TextFormat As Variant)
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
  Set qt = ws.QueryTables.add(Connection:="TEXT;" & FilePath, Destination:=ws.Range("A" & endLine))
  With qt
    If encode = "sjis" Then
      .TextFilePlatform = 932
    ElseIf encode = "utf-8" Then
      .TextFilePlatform = 65001
    End If
    .TextFileParseType = xlDelimited ' 文字で区切った形式
    .TextFileCommaDelimiter = True   ' 区切り文字はカンマ
    .TextFileStartRow = readLine     ' 1行目から読み込み
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

'==================================================================================================
' * ファイルインポート
Function importText(FilePath As String, Optional encode As String = "sjis")
  Dim buf As String, tmp As Variant, tmpJ As Variant, i As Long, j As Long

  With CreateObject("ADODB.Stream")
    .Charset = encode
    .Open
    .LoadFromFile FilePath
    buf = .ReadText
    .Close
  End With
  tmp = Split(buf, vbLf)
  For i = 0 To UBound(tmp)
    j = 0
    For Each tmpJ In Split(tmp(i), ",")
      Cells(i + 1, j + 1) = tmpJ
      j = j + 1
    Next
  Next
End Function

'==================================================================================================
Function importXlsx(FilePath As String, targeSheet As String, targeArea As String, dictSheet As Worksheet, Optional passWord As String)

  On Error GoTo catchError
  If passWord <> "" Then
    Workbooks.Open fileName:=FilePath, ReadOnly:=True, passWord:=passWord
  Else
    Workbooks.Open fileName:=FilePath, ReadOnly:=True
  End If

  If Worksheets(targeSheet).Visible = False Then
    Worksheets(targeSheet).Visible = True
  End If
  Sheets(targeSheet).Select

  ActiveWorkbook.Sheets(targeSheet).Rows.Hidden = False
  ActiveWorkbook.Sheets(targeSheet).Columns.Hidden = False

  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

  ActiveWorkbook.Sheets(targeSheet).Range(targeArea).copy
  dictSheet.Range("A1").PasteSpecial xlPasteValues

  Application.CutCopyMode = False
  ActiveWorkbook.Close SaveChanges:=False
  dictSheet.Range("A1").Select

  DoEvents
  Call Library.startScript

    Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * パスワード生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String
  Dim i As Integer, n

  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
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
  Dim f   As Font
  Dim i, iLen
  Dim r   As Range

  iLen = Len(a_sSearch)
  i = 1

  For Each r In Selection
    Do
      i = InStr(i, r.Value, a_sSearch)
      If (i = 0) Then
        i = 1
        Exit Do
      End If
      Set f = r.Characters(i, iLen).Font
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
  Const funcName As String = "Library.setRegistry"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
'  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
'  Call Library.showDebugForm("MainKey", thisAppName, "debug")
'  Call Library.showDebugForm("SubKey ", RegistrySubKey, "debug")
'  Call Library.showDebugForm("Key    ", RegistryKey, "debug")
'  Call Library.showDebugForm("Val    ", CStr(setVal), "debug")

  Call SaveSetting(thisAppName, RegistrySubKey, RegistryKey, setVal)
  
'  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function getRegistry(RegistryKey As String, RegistrySubKey As String, Optional typeVal As String = "String")
  Dim regVal As String
  Const funcName As String = "Library.getRegistry"

  On Error GoTo catchError
'  Call Library.showDebugForm(funcName, , "start1")
  
  If RegistryKey <> "" Then
    regVal = GetSetting(thisAppName, RegistryKey, RegistrySubKey)
  End If
  
'  Call Library.showDebugForm("MainKey", thisAppName, "debug")
'  Call Library.showDebugForm("Key    ", RegistryKey, "debug")
'  Call Library.showDebugForm("SubKey ", RegistrySubKey, "debug")
'  Call Library.showDebugForm("Val    ", regVal, "debug")
'  Call Library.showDebugForm("type   ", typeVal, "debug")
  
  Select Case typeVal
    Case "Boolean", "Long"
      If regVal = "" Then
        getRegistry = 0
      Else
        getRegistry = regVal
      End If
      
    Case "String", "string"
      getRegistry = regVal
    Case Else
  End Select
  
'  Call Library.showDebugForm(funcName, , "end1")
  Exit Function

'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function delRegistry(RegistryKey As String, Optional RegistrySubKey As String = "")
  Dim regVal As String

  Const funcName As String = "Library.delRegistry"
  On Error GoTo catchError
  'Call Library.showDebugForm(funcName, , "function")

  If RegistrySubKey = "" Then
    Call DeleteSetting(thisAppName, RegistryKey)
  Else
    Call DeleteSetting(thisAppName, RegistryKey, RegistrySubKey)
  End If
  Exit Function

'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * シートの保護/保護解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setProtectSheet(Optional thisAppPasswd As String)
  Const funcName As String = "Library.setProtectSheet"
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  ActiveSheet.Protect passWord:=thisAppPasswd, DrawingObjects:=True, Contents:=True, Scenarios:=True
  ActiveSheet.EnableSelection = xlUnlockedCells

  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function unsetProtectSheet(Optional thisAppPasswd As String)
  Const funcName As String = "Library.unsetProtectSheet"
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  ActiveSheet.Unprotect passWord:=thisAppPasswd
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
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
  Const funcName As String = "Library.setValandRange"

'  Range(keyName) = val
  If dicVal Is Nothing Then
    Call init.setting
  Else
    dicVal(keyName) = val
  End If
  Call Library.showDebugForm(funcName, keyName & "/" & val, "info")
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

  Call Library.showDebugForm("実行ファイル", fileName, "info")
  Call Library.showDebugForm("戻り値      ", rtnVal, "info")

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

  Const funcName As String = "Library.replaceFromFile"
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

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
  Set Txt = Nothing
  Set FSO = Nothing
  Exit Function

'エラー発生時------------------------------------
catchError:
  FSO.DeleteFile fileName & "_"
  Set Txt = Nothing
  Set FSO = Nothing
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * VBAでExcelのコメントを一括で自動サイズにしてカッコよくする
' *
' * @Link   http://techoh.net/customize-excel-comment-by-vba/
'**************************************************************************************************
Function setComment(Optional BgColorVal = 14811135, Optional FontVal = "MS UI Gothic", Optional FontColorVal = 0, Optional FontSizeVal = 9)
  Dim cl As Range
  Dim count As Long

  count = 0
  For Each cl In Selection
    count = count + 1
    DoEvents
    If Not cl.Comment Is Nothing Then
      With cl.Comment.Shape
        'サイズ設定
        .TextFrame.AutoSize = True
        .TextFrame.Characters.Font.Size = FontSizeVal
        .TextFrame.Characters.Font.Color = FontColorVal

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
' * リンク解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLink()
  Dim Links As Variant
  Dim i As Integer

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
Function 罫線_クリア_中央線_横(Optional SetArea As Range)
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
Function 罫線_クリア_中央線_縦(Optional SetArea As Range)
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


'==================================================================================================
Function 罫線_表(Optional SetArea As Range, Optional LineColor As Variant)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_囲み(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_格子(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_左(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_左右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_上(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_上下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_垂直(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_破線_水平(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_囲み(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_格子(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_左(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_左右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_上(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_上下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_垂直(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_実線_水平(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_囲み(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_左(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

  If TypeName(SetArea) = "Range" Then
    With SetArea
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function



'==================================================================================================
Function 罫線_二重線_左右(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_上(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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
Function 罫線_二重線_上下(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = dicVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)

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

'**************************************************************************************************
' * カラム幅設定 / 取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getColumnWidth()
  Dim slctCells As Range

  For Each slctCells In Selection
    slctCells = slctCells.ColumnWidth
    slctCells.HorizontalAlignment = xlCenter
    slctCells.VerticalAlignment = xlCenter
  Next
End Function

'==================================================================================================
Function setColumnWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  Const funcName As String = "Library.setColumnWidth"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  endColLine = Cells(1, Columns.count).End(xlToLeft).Column

  For colLine = 1 To endColLine
    If IsNumeric(Cells(1, colLine)) Then
      Cells(1, colLine).ColumnWidth = Cells(1, colLine)
      End If
  Next
  Exit Function
'エラー発生時------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * ページのステータス確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getURLStatusCode(ByVal strURL As String) As Integer
  'Dim Http As New WinHttpRequest
  Dim Http As Object
  Dim statusCode As Integer
  Const funcName As String = "Library.getURLStatusCode"

  On Error GoTo catchError
  Call Library.showDebugForm("URL", strURL, "info")
  If strURL = "" Then
    Exit Function
  End If
  Call Library.showDebugForm(funcName, , "start1")
  Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

  With Http
    .Open "GET", strURL, False

    If dicVal("proxyURL") <> "" Then
      .SetProxy 2, dicVal("proxyURL") & ":" & dicVal("proxyPort")
    End If
    If dicVal("proxyUser") <> "" Then
      .setProxyCredentials dicVal("proxyUser"), dicVal("proxyPasswd")
    End If

    .Send
    Call Library.showDebugForm("Status", .Status, "info")
    If .Status = 301 Or .Status = 302 Then
      Call Library.showDebugForm("GetAllResponseHeaders", .GetAllResponseHeaders, "debug")
      statusCode = .Status
    Else
      statusCode = .Status
    End If
  End With
  getURLStatusCode = statusCode


  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & Err.Description & ">", True)
  getURLStatusCode = 404
  Set Http = Nothing
End Function


'**************************************************************************************************
' * シート保護/解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function sheetProtect(Optional mode As String = "")
  Dim cellAddres As String
  Dim sheetName As Variant
  Const funcName As String = "Library.sheetProtect"

  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("mode", mode, "debug")
  Call init.setting

  If mode = "all" Then
    For Each sheetName In Sheets
      ThisWorkbook.Worksheets(sheetName.Name).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, passWord:=thisAppPasswd
      ThisWorkbook.Worksheets(sheetName.Name).EnableSelection = xlUnlockedCells

      Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      DoEvents
    Next

  ElseIf mode = "ExcelHelp" Then
    For Each sheetName In Sheets
      If sheetName.Name Like "《*》" Then
      Else
        ThisWorkbook.Worksheets(sheetName.Name).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, passWord:=thisAppPasswd
        ThisWorkbook.Worksheets(sheetName.Name).EnableSelection = xlUnlockedCells

        Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      End If
      DoEvents
    Next

  ElseIf mode = "" Then
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, passWord:=thisAppPasswd
    ActiveSheet.EnableSelection = xlUnlockedCells

    Call Library.showDebugForm("sheetName", ActiveSheet.Name, "info")
  End If
End Function

'==================================================================================================
Function sheetUnprotect(Optional allSheetflg As Boolean = False)
  Dim sheetName As Variant
  Const funcName As String = "Library.sheetUnprotect"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------

  Call Library.showDebugForm("allSheetflg", allSheetflg, "debug")
  If allSheetflg = True Then
    For Each sheetName In Sheets
      If sheetName.Name Like "《*》" Then
      Else
        ThisWorkbook.Worksheets(sheetName.Name).Unprotect passWord:=thisAppPasswd
      End If
      DoEvents
    Next
  Else
    ActiveSheet.Unprotect passWord:=thisAppPasswd
  End If

  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & Err.Description & ">", True)
  Call Library.errorHandle
End Function


'**************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function sheetNoDisplay(Optional mode As String = "")
  Dim cellAddres As String
  Dim sheetName As Variant
  Const funcName As String = "Library.sheetProtect"

  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("mode", mode, "debug")
  Call init.setting

  If mode = "all" Then
    For Each sheetName In Sheets
      ThisWorkbook.Worksheets(sheetName.Name).Visible = xlSheetVeryHidden
      Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      DoEvents
    Next

  ElseIf mode = "ehelp" Then
    For Each sheetName In Sheets
      If sheetName.Name Like "《*》" Then
        ThisWorkbook.Worksheets(sheetName.Name).Visible = xlSheetVeryHidden
        Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      End If
      DoEvents
    Next

  ElseIf mode = "" Then
    ActiveSheet.Visible = xlSheetVeryHidden
    Call Library.showDebugForm("sheetName", ActiveSheet.Name, "info")
  End If

End Function

'==================================================================================================
Function sheetDisplay(Optional mode As String = "")
  Dim cellAddres As String
  Dim sheetName As Variant
  Const funcName As String = "Library.sheetProtect"

  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("mode", mode, "info")
  Call init.setting

  If mode = "all" Then
    For Each sheetName In Sheets
      ThisWorkbook.Worksheets(sheetName.Name).Visible = True
      Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      DoEvents
    Next

  ElseIf mode = "ehelp" Then
    For Each sheetName In Sheets
      If sheetName.Name Like "《*》" Then
        ThisWorkbook.Worksheets(sheetName.Name).Visible = True
        Call Library.showDebugForm("sheetName", sheetName.Name, "info")
      End If
      DoEvents
    Next

  Else
    ThisWorkbook.Worksheets(mode).Visible = True
  End If
End Function

'**************************************************************************************************
' * 文字列連結
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setTextJoin(delimiter As String, ParamArray vals())
  Dim i As Integer
  Dim retVal As String
  Const funcName As String = "Library.setTextJoin"

  For i = LBound(vals) To UBound(vals)
    If retVal = "" Then
      retVal = vals(i)
    Else
      retVal = retVal & delimiter & vals(i)
    End If
  Next

  Call Library.showDebugForm(funcName, retVal)
  setTextJoin = retVal
End Function

'**************************************************************************************************
' * 2次元配列の1次元目をRedim Preserveする
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function RedimPreserve2D(ByVal orgArray, ByVal lengthTo)
  Dim transedArray()

  transedArray = WorksheetFunction.Transpose(orgArray)
  ReDim Preserve transedArray(1 To UBound(transedArray, 1), 1 To lengthTo)
  RedimPreserve2D = WorksheetFunction.Transpose(transedArray)
End Function



'==================================================================================================
'コントロールパネルのホイール量取得
Function getScrollRow()
  Dim scrollVal As Long
  Const GetWheelScrollLines = 104
  Const funcName As String = "Library.getScrollRow"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  SystemParametersInfo GetWheelScrollLines, 0, scrollVal, 0
  Call Library.setValandRange("scrollRowCnt", CStr(scrollVal))

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'コントロールパネルのホイール量設定
Function setScroll(setScrollRow As Long)
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
  Const funcName As String = "Library.setScroll"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.showDebugForm("setScrollRow", setScrollRow, "debug")
  '----------------------------------------------

  SystemParametersInfo SetWheelScrollLines, setScrollRow, 0, SENDCHANGE


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * iniファイル読み込み
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getConfigIni(FilePath As String)
  Dim buf As String
  Dim SectionVal As String, keyName As String, keyVal As String
  Dim keys As Variant

  Const funcName As String = "Library.getConfig"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  Set setIni = Nothing
  Set setIni = CreateObject("Scripting.Dictionary")


  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    .LoadFromFile FilePath
    Do Until .EOS
      buf = .ReadText(-2)

      If Len(buf) = 0 Then
      ElseIf Left(buf, 1) = ";" Then
      ElseIf Left(buf, 1) = "[" Then
        SectionVal = Mid(buf, 2, Len(buf) - 2)
      ElseIf InStr(1, buf, "=") > 0 Then
        keys = Split(buf, "=")
        keyName = keys(0)
        keyVal = keys(1)

        Call Library.showDebugForm(SectionVal & "_" & keyName, keyVal, "debug")
        setIni.add SectionVal & "_" & keyName, keyVal
      End If
    Loop
    .Close
  End With

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
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
Function setConfigIni(sectionName As String, keyName As String, FilePath As String, setVal As String)
  Const funcName As String = "Library.setLineHeight"


  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

'  Call WritePrivateProfileString(sectionName, keyName, setVal, filePath)



  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function getLineHeight(targetRang As Range, maxLen As Long, defaultRowHeight As Long)
  Dim LFCount As Long, LenCount As Long
'  Dim setHeight As Long
  Const funcName As String = "Library.getLineHeight"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  LFCount = UBound(Split(targetRang.Value, vbNewLine))
  LenCount = Library.getLength(targetRang.Value)

  If LFCount > 0 Then
    LFCount = LFCount + 1
  Else
    If LenCount > maxLen Then
      LFCount = Int(LenCount / maxLen) + 1
    Else
      LFCount = 1
    End If
  End If
  Call Library.showDebugForm("LFCount", LFCount, "debug")
  Call Library.showDebugForm("LenCount", LenCount, "debug")

  getLineHeight = LFCount

'  setHeight = defaultRowHeight * LFCount
'
'  Call Library.showDebugForm("LFCount", LFCount, "debug")
'  Call Library.showDebugForm("LenCount", LenCount, "debug")
'  Call Library.showDebugForm("setHeight", setHeight, "debug")
'
'  If ActiveSheet.Rows(targetRang.Row & ":" & targetRang.Row).RowHeight < setHeight Then
'    ActiveSheet.Rows(targetRang.Row & ":" & targetRang.Row).RowHeight = setHeight
'  End If

  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function getLineWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  Dim slctCells As Range
  Const funcName As String = "Library.getLineWidth"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------

  Cells.EntireColumn.AutoFit

  For colLine = 1 To Columns.count
    If Cells(1, colLine).ColumnWidth > 30 Then
      colName = Library.getColumnName(colLine)
      Columns(colName & ":" & colName).ColumnWidth = 30
    End If
  Next

  '処理終了--------------------------------------

  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function





'==================================================================================================
' * 文字列の左側から指定文字数抽出
Function getLeftString(targetStr, strCnt As Long) As String
  Dim targetLen As Long
  Dim getStr As String

  targetLen = Len(targetStr)

  '文字列ではない場合
  If VarType(targetStr) <> vbString Then
    getStr = targetStr

  ElseIf targetLen < strCnt Then
    getStr = targetStr
  Else
    'getStr = Right(targetStr, targetLen - strCnt)
    getStr = Left(targetStr, strCnt)
  End If

  getLeftString = getStr

End Function

'==================================================================================================
' * 文字列の右側から指定文字数抽出
Function getRightString(targetStr, strCnt As Long) As String
  Dim targetLen As Long
  Dim getStr As String

  targetLen = Len(targetStr)

  If VarType(targetStr) <> vbString Then
    getStr = targetStr

  ElseIf targetLen < strCnt Then
    getStr = targetStr
  Else
    getStr = Right(targetStr, strCnt)
  End If

  getRightString = getStr
End Function


'==================================================================================================
' * クイックソート
Sub Sort_QuickSort(ByRef argAry As Variant, ByVal lngMin As Long, ByVal lngMax As Long, ByVal keyPos As Long)
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim vBase As Variant
  Dim vSwap As Variant
  
  vBase = argAry(Int((lngMin + lngMax) / 2), keyPos)
  i = lngMin
  j = lngMax
  Do
    Do While argAry(i, keyPos) < vBase
      i = i + 1
    Loop
    Do While argAry(j, keyPos) > vBase
      j = j - 1
    Loop
    If i >= j Then Exit Do
    For k = LBound(argAry, 2) To UBound(argAry, 2)
      vSwap = argAry(i, k)
      argAry(i, k) = argAry(j, k)
      argAry(j, k) = vSwap
    Next
    i = i + 1
    j = j - 1
  Loop
  If (lngMin < i - 1) Then
    Call Library.Sort_QuickSort(argAry, lngMin, i - 1, keyPos)
  End If
  If (lngMax > j + 1) Then
    Call Library.Sort_QuickSort(argAry, j + 1, lngMax, keyPos)
  End If
End Sub


'==================================================================================================
Function Bookの状態確認() As Boolean
  Dim wb As Workbook, tmp As String
  Dim retFlg As Boolean
  
  Const funcName As String = "Library.Bookの状態確認"
  
  '処理開始--------------------------------------
  On Error Resume Next
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  retFlg = False
  
  '開いているBookの数を確認----------------------
  If Workbooks.count = 0 Then
    Call Library.showDebugForm("ブックが開かれていません", , "Error")
    retFlg = False
  Else
    Call Library.showDebugForm("Workbooks.count", Workbooks.count, "debug")
    retFlg = True
  End If
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("retFlg", retFlg, "debug")
  Call Library.showDebugForm(funcName, , "end1")
  Bookの状態確認 = retFlg
  Exit Function
  '----------------------------------------------
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function スタイル利用確認()
  Dim count As Long, endCount As Long
  Dim i As Long, RangeCnt As Long
  Dim objSheet As Variant
  Dim sheetName As String, styleName As String
  Dim slctRange As Range
  
  Const funcName As String = "Library.スタイル利用確認"
  
  '処理開始--------------------------------------
  On Error Resume Next
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  ReDim useStyle(0)
  useStyle(0) = "標準"
  
  i = 1
  RangeCnt = 1
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    
    For Each slctRange In Worksheets(sheetName).UsedRange
      styleName = slctRange.style.NameLocal
      
      If Library.chkArrayVal(useStyle, styleName) = False Then
        ReDim Preserve useStyle(i)
        useStyle(i) = styleName
        i = i + 1
        DoEvents
      End If
      
      'Call Ctl_ProgressBar.showBar("スタイル利用確認", PrgP_Cnt, PrgP_Max, RangeCnt, Worksheets(sheetName).UsedRange.count, "シート：" & sheetName)
      RangeCnt = RangeCnt + 1
    Next
    DoEvents
  Next


  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function スタイル削除()
  Dim objStyle As Variant
  Dim count As Long, endCount As Long
  Dim retFlg As Boolean
  
  Const funcName As String = "Library.スタイル削除"
  
  '処理開始--------------------------------------
  On Error Resume Next
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  count = 1
  endCount = ActiveWorkbook.Styles.count
  
  For Each objStyle In ActiveWorkbook.Styles
    Call Ctl_ProgressBar.showBar("定義済スタイル削除", PrgP_Cnt, PrgP_Max, count, endCount, "シート：" & objStyle.Name)
    
    Call Library.showDebugForm("スタイル      ", objStyle.Name, "debug")
    If Library.chkArrayVal(useStyle, objStyle.Name) = False Then
      Call Library.showDebugForm("削除スタイル  ", objStyle.Name, "debug")
      Select Case objStyle.Name
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Case Else
          objStyle.delete
      End Select
    End If
    count = count + 1
  Next
  
  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

