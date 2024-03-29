Attribute VB_Name = "Library"
Option Explicit

'**************************************************************************************************
' * QÆÝèAèé¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' Windows APIÌp--------------------------------------------------------------------------------
#If VBA7 And Win64 Then
  'fBXvCÌðxæ¾p
  Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'SleepÖÌp
  Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As LongPtr)

  'Nbv{[hÖA
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

#Else
  'fBXvCÌðxæ¾p
  Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'SleepÖÌp
  Private Declare Function Sleep Lib "kernel32" (ByVal ms As Long)

  'Nbv{[hÖA
  Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function EmptyClipboard Lib "user32" () As Long

  'ShellÖÅN®µ½vOÌI¹ðÒÂ
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Private Const PROCESS_QUERY_INFORMATION = &H400&
  Private Const STILL_ACTIVE = &H103&
#End If


'[NubNpÏ------------------------------
'[NV[gpÏ------------------------------
'O[oÏ----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

Public CalculatFlg As Boolean


'ANeBuZÌæ¾
Dim SelectionCell As String
Dim SelectionSheet As String

' PCAOfficeÌîñæ¾pAzzñ
Public MachineInfo As Object

'_Öp
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

Public Const JapaneseCharacters = " ¢¤¦¨©«­¯±³µ·¹»½¿ÂÄÆÈÉÊËÌÍÐÓÖÙÜÝÞßàâäæçèéêëíðñª¬®°²´¶¸º¼¾ÀÃÅÇÎÑÔ×ÚÏÒÕØÛ"
Public Const JapaneseCharactersCommonUse = "JwóàÂÑæâom·¼Xà¾åéÏçjÝêïKngÀÒÌæóègÂ\¨ú¡½ûapÊè¯G¦aÅüQ¡ü¼²êITÅO{@qáÕ¿ÍøÈ}Éq³µ«§ÅìñgÄá@SÍûqÀóËïFµü¨hÀª©ñïËtÞ|øbÎ¢£"
Public Const MachineDependentCharacters = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]_ú\ú]ú^ú_ú`úaúbúcúdúeúfúgúhúiújúkúlúmúnúoúpúqúrúsútúuúvúwúxúyúzú{"

Public ThisBook As Workbook


'**************************************************************************************************
' * G[Ì
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
' * æÊ`Ê§äJn
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()
  Const funcName As String = "Library.startScript"

  On Error Resume Next
  If logFile <> "" Then
    Call Library.showDebugForm(funcName, , "function1")
  End If
  
  'ANeBuZÌæ¾
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
    SelectionSheet = ActiveWorkbook.ActiveSheet.Name
  End If

  '}N®ìÅZItemNameÈÇªÏíé©®vZªðx­·éÌðð¯é
  If Application.Calculation = xlCalculationManual Then
    CalculatFlg = False
  Else
    Application.Calculation = xlCalculationManual
    CalculatFlg = True
  End If
  Call Library.showDebugForm("CalculatFlg", CalculatFlg, "debug")

  Application.ScreenUpdating = False              '}N®ìÅV[gâEBhEªØèÖíéÌð©¹È¢æ¤ÉµÜ·
  Application.EnableEvents = False                '}N®ì©ÌÅÊÌCxgª¶¬³êéÌð}§·é
  
  Application.DisplayAlerts = False               'mFbZ[Wðo³È¢
  'Application.StatusBar = "EEE"         'Xe[^Xo[Éð\¦

'  If runFlg = True Then
'    Application.Interactive = False                 '}N®ìÉêØÌL[â}EXìð§À·é
'    Application.Cursor = xlWait                     '}N®ìÍ}EXJ[\ðu»vvÉ·é
'  End If

End Function

'**************************************************************************************************
' * æÊ`Ê§äI¹
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional reCalflg As Boolean = False, Optional flg As Boolean = False)
  Const funcName As String = "Library.endScript"

  On Error Resume Next
  If logFile <> "" Then
    Call Library.showDebugForm(funcName, , "function1")
  End If

  '­§IÉÄvZ³¹é
  If reCalflg = True Then
    Application.CalculateFull
  Else
    ActiveSheet.Calculate
  End If
  
  Call Library.showDebugForm("CalculatFlg", CalculatFlg, "debug")
  If CalculatFlg = True Then
    Application.Calculation = xlCalculationAutomatic  '}N®ìÅZItemNameÈÇªÏíé©®vZªðx­·éÌðð¯é
  End If

 'ANeBuZÌIð
  If SelectionCell <> "" And flg = True Then
    ActiveWorkbook.Worksheets(SelectionSheet).Select
    ActiveWorkbook.Range(SelectionCell).Select
  End If

  Application.ScreenUpdating = True                 '}N®ìÅV[gâEBhEªØèÖíéÌð©¹È¢æ¤ÉµÜ·
  Application.EnableEvents = True                   '}N®ì©ÌÅÊÌCxgª¶¬³êéÌð}§·é
  
  Application.Interactive = True                    '}N®ìÉêØÌL[â}EXìð§À·é
  Application.Cursor = xlDefault                    '}N®ìI¹ãÍ}EXJ[\ðuftHgvÉàÇ·
  Application.StatusBar = False                     '}N®ìI¹ãÍXe[^Xo[ðuftHgvÉàÇ·
  Application.DisplayAlerts = True                  'mFbZ[Wðo³È¢
End Function

'**************************************************************************************************
' * V[gÌ¶ÝmF
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
' * ®¹ÜÅÒ@
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
' * I[gVFCvÌ¶ÝmF
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
' * OV[g»è
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
' * zñªó©Ç¤©
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
'G[­¶------------------------------------
catchError:
  chkArrayEmpty = True
End Function

'**************************************************************************************************
' * zñÉlª¶Ý·é©Ç¤©
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
'G[­¶------------------------------------
catchError:
  chkArrayVal = True
End Function

'**************************************************************************************************
' * ubNªJ©êÄ¢é©`FbN
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
' * wb_[`FbN
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
    errMeg = "ÂªÙÈèÜ·B"
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
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'**************************************************************************************************
' * f[^`FbN
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'út
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
Function chkIsOpen(targetBookName As String) As Boolean
  Dim openWorkbook As Workbook
  Dim chkFlg As Boolean
  Const funcName As String = "Library.chkIsOpen"

  chkFlg = False
  For Each openWorkbook In Workbooks
    If openWorkbook.Name = targetBookName Then
      chkFlg = True
      Exit For
    End If
  Next
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("targetBookName", targetBookName, "debug")
  Call Library.showDebugForm("isOpen", chkFlg, "debug")
  Call Library.showDebugForm(funcName, , "end1")
  
  chkIsOpen = chkFlg
End Function


'**************************************************************************************************
' * t@CÌÛ¶êª[JfBXN©Ç¤©»è
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

  'hCuÌíÞð»Ê
  If driveName = "" Then
    driveType = 0 's¾
  Else
    driveType = FSO.GetDrive(driveName).driveType
  End If

  Select Case driveType
    Case 1
      retVal = True
      Call Library.showDebugForm("Library.chkLocalDrive", "[oufBXN")
    Case 2
      retVal = True
      Call Library.showDebugForm("Library.chkLocalDrive", "n[hfBXN")
    Case Else
      retVal = False
      Call Library.showDebugForm("Library.chkLocalDrive", "s¾Albg[NhCuACDhCuÈÇ")
  End Select

  If LadexsetVal("debugMode") = "develop" Then
    retVal = False
  End If
  chkLocalDrive = retVal
  Exit Function
'G[­¶------------------------------------
catchError:
End Function


'**************************************************************************************************
' * pX©çt@C©fBNg©ð»è
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
' * t@CÌ¶ÝmF
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
' * fBNgÌ¶ÝmF
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
' * lÌ^`FbN
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkTypeName(targetVal As Variant, permitType As String, Optional regPattern As String)
  Dim regexp
  Dim resultFlg As Boolean

  Const funcName As String = "Library.chkTypeName"

  'Jn--------------------------------------
  On Error GoTo catchError
  Set regexp = CreateObject("VBScript.RegExp")
  resultFlg = False

  If targetVal = "" Then
    resultFlg = True
    GoTo Lbl_endSelect
  End If
  regexp.Global = True              '¶ñSÌðõ

  Select Case permitType
    Case "int"        '
      regexp.Pattern = "^[0-9]+$"
      resultFlg = regexp.test(targetVal)

    Case "string"     'p¬¶
      regexp.IgnoreCase = False
      regexp.Pattern = "^[a-z]+$"
      resultFlg = regexp.test(targetVal)

    Case "STRING"     'på¶
      regexp.IgnoreCase = False
      regexp.Pattern = "^[A-Z]+$"
      resultFlg = regexp.test(targetVal)

    Case "String"     'p(å¬æÊÈµ)
      regexp.IgnoreCase = True
      regexp.Pattern = "^[a-zA-Z]+$"
      resultFlg = regexp.test(targetVal)

    Case "reg"        '³K\»
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
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * Byte©çKB,MB,GBÖÏ·
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
' * Åè·¶ñÉÏ·
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
' * LP[XðXl[NP[XÉÏ·
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
' * Xl[NP[XðLP[XÉÏ·
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
' * ¼pðSpÉÏ··é(pAJ^Ji)
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
' * Spð¼pÉÏ··é(pAJ^Ji)
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
    If buf Like "[`-O-X]" Or _
      buf Like "[|IDopij^]" Then
      covText = covText & StrConv(buf, vbNarrow)
        
    ElseIf buf Like "[±-Ý]" Then
      covText = covText & StrConv(buf, vbWide)
    
    ElseIf buf = "," Then
      covText = covText & "C"
    
    Else
      covText = covText & buf
    End If
    DoEvents
  Next i
  
  Call Library.showDebugForm(funcName, covText, "debug")
  convZen2Han = covText
End Function

'**************************************************************************************************
' * pCvðJ}ÉÏ·
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
' * Base64GR[h(t@C)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" 'ú»
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
' * Base64GR[h(¶ñ)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String
  Dim ret As String
  Dim d() As Byte
  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" 'ú»
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
' * URL-safe Base64GR[h
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
' * URLGR[h
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLEncode(ByVal strVal As String) As String
'  Dim EncodeURL As String

'  With CreateObject("ScriptControl")
'    .Language = "JScript"
'    EncodeURL = .CodeObject.encodeURIComponent(strVal)
'  End With
'  convURLEncode = EncodeURL
  
  convURLEncode = WorksheetFunction.EncodeURL(strVal)
End Function


'**************************************************************************************************
' * URLfR[h
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLDecode(ByVal strVal As String) As String
  Dim DecodeURL As String

  With CreateObject("ScriptControl")
    .Language = "JScript"
    DecodeURL = .CodeObject.decodeURIComponent(strVal)
  End With
  convURLDecode = DecodeURL
End Function


'**************************************************************************************************
' * UnicodeGXP[v
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
    Call Library.showDebugForm("ÎÛ¶ñ", Mid$(escape, i, 1) & "<:>" & acode, "debug")
    
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
' * UnicodeGXP[v
' *
' * @link   http://tech7.blog.shinobi.jp/vba/unicode%E3%82%A8%E3%82%B9%E3%82%B1%E3%83%BC%E3%83%97%E3%81%95%E3%82%8C%E3%81%9F%E6%96%87%E5%AD%97%E5%88%97%E3%82%92%E6%96%87%E5%AD%97%E3%81%AB%E6%88%BB%E3%81%99%E6%96%B9%E6%B3%95
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
' * æªP¶Úðå¶»
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
' * ¶ñÌ¶¤©çwè¶í·éÖ
' *
' * @Link   https://vbabeginner.net/vbaÅ¶ñÌE¤â¶¤©çwè¶í·é/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
  Dim iLen    As Long

  '¶ñÅÍÈ¢ê
  If VarType(s) <> vbString Then
    cutLeft = s & "¶ñÅÍÈ¢"
    Exit Function
  End If
  iLen = Len(s)
  '¶ñ·æèwè¶ªå«¢ê
  If iLen < i Then
    cutLeft = s & "¶ñ·æèwè¶ªå«¢"
    Exit Function
  End If
  cutLeft = Right(s, iLen - i)
End Function

'**************************************************************************************************
' * ¶ñÌE¤©çwè¶í·éÖ
' *
' * @Link   https://vbabeginner.net/vbaÅ¶ñÌE¤â¶¤©çwè¶í·é/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
  Dim iLen    As Long

  If VarType(s) <> vbString Then
    cutRight = s & "¶ñÅÍÈ¢"
    Exit Function
  End If
  iLen = Len(s)
  '¶ñ·æèwè¶ªå«¢ê
  If iLen < i Then
    cutRight = s & "¶ñ·æèwè¶ªå«¢"
    Exit Function
  End If
  cutRight = Left(s, iLen - i)
End Function

'**************************************************************************************************
' * A±üsÌí
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
' * V[gí
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
    targetSheet.Rows(line & ":" & Rows.count).NumberFormatLocal = "G/W"
    targetSheet.Rows(line & ":" & Rows.count).style = "Normal"
  Else
    targetSheet.Cells.delete Shift:=xlUp
    targetSheet.Cells.NumberFormatLocal = "G/W"
    targetSheet.Cells.style = "Normal"
  End If
  DoEvents

  If delImgFlg = True Then
    For Each shp In ActiveSheet.Shapes
    shp.Select
      If shp.Type = 11 Then shp.delete
    Next shp
  End If
  
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm(funcName, , "end1")
  
'  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 1, "f[^ÁF" & targetSheet.name)
  
End Function
'**************************************************************************************************
' * ZàÌüsí
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
' * IðÍÍÌæí
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
' * ZÌ¼ÌÝèí
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
' * e[uf[^í
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
  Selection.NumberFormatLocal = "G/W"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * t@CRs[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execCopy(srcPath As String, dstPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execCopy"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call showDebugForm("  Rs[³F" & srcPath)
  Call showDebugForm("  Rs[æF" & dstPath)

  If chkFileExists(srcPath) = False Then
    Call showinfo(404, "Rs[³", True)
  End If

  If chkDirExists(getParentDir(dstPath)) = False Then
    Call Library.execMkdir(getParentDir(dstPath))
  End If
  FSO.CopyFile srcPath, dstPath
  Set FSO = Nothing
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * t@CÚ®
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execMove(srcPath As String, dstPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execMove"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call showDebugForm("  Ú®³F" & srcPath)
  Call showDebugForm("  Ú®æF" & dstPath)

  If chkFileExists(srcPath) = False Then
    Call showinfo(404, "Ú®³", True)
  End If

  FSO.MoveFile srcPath, dstPath
  Set FSO = Nothing
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * fBNgí
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execDeldir(srcPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execDeldir"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call showDebugForm("  íÎÛF" & srcPath)

  If srcPath Like "*[*]*" Then
  ElseIf chkDirExists(srcPath) = False Then
    Call showinfo(404, "íÎÛ", True)
  End If

  FSO.DeleteFolder srcPath
  Set FSO = Nothing
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * t@Cí
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execDel(srcPath As String)
  Dim FSO As Object
  Const funcName As String = "Library.execDel"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Call showDebugForm("íÎÛ", srcPath, "debug")

  If srcPath Like "*[*]*" Then
  ElseIf chkFileExists(srcPath) = False Then
    Call showNotice(404, "íÎÛ", True)
  End If

  FSO.DeleteFile srcPath
  Set FSO = Nothing
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * t@C¼ÏX
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
  Call showDebugForm("ÏX³", srcPath)
  Call showDebugForm("¼Ì", oldFileName)
  Call showDebugForm("V¼Ì", fileName)

  If chkFileExists(srcPath & "\" & oldFileName) = False Then
    If IsMissing(errMeg) Then
      Call showinfo(404, "ÏX³", True)
    Else
      errMeg = "ÏX³Ìt@Cª èÜ¹ñ[" & oldFileName & "]"
      errFlg = True
    End If

  End If
  If chkFileExists(srcPath & "\" & fileName) = True Then
    If IsMissing(errMeg) Then
      Call showinfo(414, fileName, True)
    Else
      errMeg = "¯¼Ìt@Cª¶ÝµÜ·[" & fileName & "]"
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
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  errMeg = Err.Description
  execRename = False
End Function

'**************************************************************************************************
' * MkDirÅKwÌ[¢tH_[ðìé
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execMkdir(fullPath As String)

  If chkDirExists(fullPath) Then
    Exit Function
  End If

  Call showDebugForm("execMkdir", fullPath, "debug")
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

'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * zip³k/ð
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function execCompress(srcPath As String, zipFilePath As String) As Boolean
  'Dim sh  As New IWshRuntimeLibrary.WshShell
  Dim Sh
  Dim ex  As WshExec
  Dim cmd As String
  Set Sh = CreateObject("WScript.Shell")
  Call showDebugForm("ÎÛfBNgF" & srcPath)
  Call showDebugForm("zipt@C     F" & zipFilePath)

  If chkDirExists(srcPath) = False Then
    Call showinfo(403, "ÎÛfBNg", True)
  End If

  '¼pXy[XðobNNH[gÅGXP[v
  srcPath = Replace(srcPath, " ", "` ")
  zipFilePath = Replace(zipFilePath, " ", "` ")

  cmd = "Compress-Archive -Path " & srcPath & " -DestinationPath " & zipFilePath & " -Force"
  Call showDebugForm("cmdF" & cmd)
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
  Call showDebugForm("zipt@C     ", zipFilePath)
  Call showDebugForm("ÎÛfBNg", dstPath)

  If chkFileExists(zipFilePath) = False Then
    Call showinfo(404, "ðÎÛ", True)
  End If
  If chkDirExists(dstPath) = False Then
    Call showinfo(403, "ðæ", True)
  End If

  '¼pXy[XðobNNH[gÅGXP[v
  zipFilePath = Replace(zipFilePath, " ", "` ")
  dstPath = Replace(dstPath, " ", "` ")

  cmd = "Expand-Archive -Path " & zipFilePath & " -DestinationPath " & dstPath & " -Force"
  Call showDebugForm("cmdF" & cmd)
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
' * PCAOfficeÌîñæ¾
' * AzzñðpµÄ¢éÌÅAMicrosoft Scripting RuntimeªK{
' * MachineInfo.Item ("Excel") ÅÄÑoµ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object
  On Error Resume Next

  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  'OSÌo[Wæ¾----------------------------
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

  'ExcelÌo[Wæ¾-------------------------
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

  'PCÌîñ--------------------------------------
  MachineInfo.add "UserName", WshNetworkObject.userName
  MachineInfo.add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.add "UserDomain", WshNetworkObject.UserDomain

  'æÊÌðxæ¾----------------------------
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
' * ¶JEg
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getLength(targetVal As String, Optional charType As String = "¼p")
  Dim inputLen As Long
  Const funcName As String = "Library.getLength"
  
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm("targetVal", targetVal, "debug")
'  Call Library.showDebugForm("charType ", charType, "debug")
'  Call Library.showDebugForm("¶   [Len]", Len(targetVal), "debug")
'  Call Library.showDebugForm("oCg[LenB]", LenB(StrConv(targetVal, vbFromUnicode)), "debug")
  
  Select Case charType
    Case "¼p", "Sp"
      inputLen = LenB(StrConv(targetVal, vbFromUnicode))
    Case "¶"
      inputLen = Len(targetVal)
  End Select
  
'  Call Library.showDebugForm("inputLen", inputLen, "debug")
  getLength = inputLen
  
  Exit Function
'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * ZÌÀWæ¾
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
  Call Library.showDebugForm("R1C1Top     F" & R1C1Top)
  Call Library.showDebugForm("R1C1Left    F" & R1C1Left)
  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("Rng.Address F" & Rng.Address)
  Call Library.showDebugForm("ActvCellTop F" & ActvCellTop)
  Call Library.showDebugForm("ActvCellLeftF" & ActvCellLeft)
End Function

'**************************************************************************************************
' * ñ¼©çñÔðßé
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long
  getColumnNo = Range(targetCell & ":" & targetCell).Column
End Function

'**************************************************************************************************
' * ñÔ©çñ¼ðßé
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String
  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function

'**************************************************************************************************
' * J[pbgð\¦µAFR[hðæ¾
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
' * tHg_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFont(FontName As String, FontSize As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Application.Dialogs(xlDialogActiveCellFont).Show FontName, "M[", FontSize
End Function

'**************************************************************************************************
' * IndentLevellæ¾
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
' * RGBlæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function

'**************************************************************************************************
' * fBNgIð_CAO\¦
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
      .title = title & "ÌêðIðµÄ­¾³¢"
    Else
      .title = "tH_[ðIðµÄ­¾³¢"
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
' * t@CÛ¶_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)
  Dim filePath As String
  Dim Result As Long
  Dim fileName As Variant

  fileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excelt@C,*.xlsx,Excel2003ÈO,*.xls,Excel}NubN,*.xlsm,·×ÄÌt@C, *.*", _
      FilterIndex:=FileTypeNo)

  If fileName <> "False" Then
    getSaveFilePath = filePath
  Else
    getSaveFilePath = ""
  End If
End Function

'**************************************************************************************************
' * t@CIð_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, fileName As String, title As String, fileType As String)
  Dim filePath As String
  Dim Result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' t@CÌíÞðÝè
    .Filters.Clear
    Select Case fileType
      Case "Excel"
            .Filters.add "ExcelubN", "*.xls; *.xlsx; *.xlsm"
      Case "txt"
        .Filters.add "eLXgt@C", "*.txt"

      Case "csv"
        .Filters.add "CSVt@C", "*.csv"

      Case "json"
        .Filters.add "JSONt@C", "*.json"

      Case "sql"
        .Filters.add "SQLt@C", "*.sql"

      Case "mdb"
        .Filters.add "Accesssf[^x[X", "*.mdb;*.accdb"

      Case "img"
        .Filters.add "C[Wt@C", "*.bmp;*.jpg;*.gif;*.png"

      Case "psd"
        .Filters.add "Photoshop Data", "*.psd"

      Case "NGCeBu"
        .Filters.add "NGCeBu", "*.jpg;*.gif;*.png;*.mp4"

      Case "mov"
        .Filters.add "®æt@C", "*.mp4"

      Case Else
        .Filters.add "·×ÄÌt@C", "*.*"
    End Select
    '.FilterIndex = FileTypeNo

    '\¦·étH_
    If chkDirExists(CurrentDirectory) = True Then
      .InitialFileName = CurrentDirectory & "\" & fileName
    Else
      .InitialFileName = ActiveWorkbook.path & "\" & fileName
    End If

    '\¦`®ÌÝè
    .InitialView = msoFileDialogViewWebView

    '_CAO {bNXÌ^CgÝè
    .title = title & "ðIðµÄ­¾³¢"

    If .Show = -1 Then
      filePath = .SelectedItems(1)
    Else
      filePath = ""
    End If
  End With
  getFilePath = filePath
End Function

'**************************************************************************************************
' * ¡t@CIð_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, title As String, fileType As String, Optional setRegPathName As String = "")
  Dim filePath() As Variant
  Dim tmpPath As String
  Dim Result As Long, i As Integer

  If setRegPathName <> "" Then
    tmpPath = Library.getRegistry("targetInfo", setRegPathName)
    If tmpPath <> "" Then
      CurrentDirectory = tmpPath
    End If
  End If
  
  With Application.FileDialog(msoFileDialogFilePicker)
    '¡IððÂ
    .AllowMultiSelect = True

    ' t@CÌíÞðÝè
    .Filters.Clear
    Select Case fileType
      Case "Excel"
        .Filters.add "ExcelubN", "*.xls; *.xlsx; *.xlsm"

      Case "txt"
        .Filters.add "eLXgt@C", "*.txt"

      Case "csv"
        .Filters.add "CSVt@C", "*.csv"

      Case "json"
        .Filters.add "JSONt@C", "*.json"

      Case "sql"
        .Filters.add "SQLt@C", "*.sql"

      Case "mdb"
        .Filters.add "Accesssf[^x[X", "*.mdb;*.accdb"

      Case "img"
        .Filters.add "C[Wt@C", "*.bmp;*.jpg;*.gif;*.png"

      Case "psd"
        .Filters.add "Photoshop Data", "*.psd"

      Case "mov"
        .Filters.add "®æt@C", "*.mp4"

      Case Else
        .Filters.add "·×ÄÌt@C", "*.*"
    End Select
    '.FilterIndex = FileTypeNo

    '\¦·étH_
    .InitialFileName = CurrentDirectory & "\"

    '\¦`®ÌÝè
    .InitialView = msoFileDialogViewWebView

    '_CAO {bNXÌ^CgÝè
    .title = title

    If .Show = -1 Then
      Call Library.setRegistry("targetInfo", setRegPathName, Library.getFileInfo(.SelectedItems(1), , "CurrentDir"))
      
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
' * fBNgàÌt@Cêæ¾
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
' * t@Cîñæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileInfo(targetFilePath As String, Optional fileInfo As Object, Optional getType As String)
  Dim FSO As Object
  Dim fileObject As Object
  Dim sp As Shape

  Call Library.showDebugForm("targetFilePath", targetFilePath, "debug")
  If Library.chkFileExists(targetFilePath) = False Then
    Exit Function
  End If
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set fileInfo = Nothing
  Set fileInfo = CreateObject("Scripting.Dictionary")
  

  'ì¬ú
  fileInfo.add "createAt", Format(FSO.GetFile(targetFilePath).DateCreated, "yyyy/mm/dd hh:nn:ss")

  'XVú
  fileInfo.add "updateAt", Format(FSO.GetFile(targetFilePath).DateLastModified, "yyyy/mm/dd hh:nn:ss")

  't@CTCY
  fileInfo.add "size", FSO.GetFile(targetFilePath).Size

  't@CÌíÞ
  fileInfo.add "type", FSO.GetFile(targetFilePath).Type

  'g£q
  fileInfo.add "extension", FSO.GetExtensionName(targetFilePath)

  't@C¼
  fileInfo.add "fileName", FSO.GetFile(targetFilePath).Name

  't@Cª¶Ý·étH_
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
' * t@CÌetH_æ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getParentDir(targetPath As String) As String
  Dim parentDir As String

  parentDir = Left(targetPath, InStrRev(targetPath, "\") - 1)
'  Call Library.showDebugForm(" parentDirF" & parentDir)
  getParentDir = parentDir
End Function

'**************************************************************************************************
' * wèoCgÌÅè·f[^ì¬(¶ñ)
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
  ' oCg»è
  For ix = 1 To lngKeta
    ' 1¶¸Â¼p/Spð»f
    intCHAR = Asc(Mid(strInText, ix, 1))
    ' SpÆ»f³êéêÍoCgÉ1ðÁ¦é
    If ((intCHAR < 0) Or (intCHAR > 255)) Then
        lngByte2 = 2        ' Sp
    Else
        lngByte2 = 1        ' ¼p
    End If
    '  Óê»è(EØèÌÄ)
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
  ' s«»è(ó¶ÇÁ)
  If lngByte < lngFixBytes Then
      strOutText = strOutText & Space(lngFixBytes - lngByte)
  End If
  getFixlng = strOutText
End Function

'**************************************************************************************************
' * V[gXgæ¾
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

  '»ÝèlÌNA
  Worksheets("Ýè").Range(columnName & "3:" & columnName & "100").Select
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
    'V[g¼ÌÝè
    Worksheets("Ýè").Range(columnName & i).Select
    Worksheets("Ýè").Range(columnName & i) = sheetName.Name

    ' ZÌwiFð
    With Worksheets("Ýè").Range(columnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' V[gFÆ¯¶FðZÉÝè
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("Ýè").Range(columnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    'rüÌÝè
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

  Worksheets("Ýè").Range(columnName & "3").Select
  Call endScript
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * IðZÌgå\¦Äo
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
' * fobOpæÊ\¦
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
      meg1 = "  [Error] " & Replace(meg1, vbNewLine, " ")
      LogLevel = 1

    Case "warning"
      meg1 = "  [warn]  " & Replace(meg1, vbNewLine, " ")
      LogLevel = 2

    Case "info"
      meg1 = "  [info]  " & Replace(meg1, vbNewLine, " ")
      LogLevel = 4

    Case "debug"
      meg1 = "  [debug] " & Replace(meg1, vbNewLine, " ")
      LogLevel = 5

    Case "start"
      meg1 = Library.convFixedLength(meg1, 62, "=")
      LogLevel = 0
    Case "end"
      meg1 = Library.convFixedLength("", 62, "=")
      LogLevel = 0
      
    Case "function"
      meg1 = meg1
      LogLevel = 0
    
    Case "start1"
      meg1 = Library.convFixedLength("  " & meg1, 60, "-")
      LogLevel = 0
    Case "end1"
      meg1 = Library.convFixedLength("  ", 60, "-")
      LogLevel = 0
    Case "function1"
      meg1 = "  " & meg1
      LogLevel = 0
      
    Case Else
      meg1 = "  [] " & Replace(meg1, vbNewLine, " ")
      LogLevel = 6
  End Select

  If IsMissing(meg2) = False Then
    meg1 = meg1 & " : " & Application.WorksheetFunction.Trim(CStr(meg2))
  End If

  If CInt(LogLevel) <= CInt(Split(LadexsetVal("LogLevel"), ".")(0)) Then
    Call outputLog(runTime, meg1)
  End If
  If LadexsetVal("debugMode") = "develop" Then
    Debug.Print runTime & "  " & meg1
  End If
  DoEvents

  If LogLevel = 6 Then
    Stop
  End If
  Exit Function

'G[­¶------------------------------------
catchError:
  
  Debug.Print runTime & "  [ERROR] " & Err.Description; "  " & meg1
  Exit Function
End Function

'**************************************************************************************************
' * Xe[^Xo[ÉbZ[Wð\¦
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

'G[­¶------------------------------------
catchError:
  Debug.Print "  [ERROR] " & Err.Description; "  " & message
  Exit Function
End Function

'**************************************************************************************************
' * îñÊm
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
    speakerMeg = message & vbNewLine & "Bð~µÜ·"
    errLevel = "Error"
  Else
    speakerMeg = message
  End If

  If message <> "" Then
    message = Replace(message, "<BR>", vbNewLine)
  End If

  If LadexsetVal("debugMode") = "speak" Or LadexsetVal("debugMode") = "develop" Or LadexsetVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=speakerMeg, SpeakAsync:=True, SpeakXML:=True
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

  'æÊ`Ê§äI¹
  If runEndflg = True Then
    Call Library.endScript
    Call Ctl_ProgressBar.showEnd
    Call init.unsetting
    End
  End If

  Exit Function
'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(message, , errLevel)
  Call MsgBox(message, vbCritical, thisAppName)
End Function

'**************************************************************************************************
' * _
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
    'WFl[^ðú»
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
' * OoÍ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function outputLog(runTime As String, message As String)
  Dim fileTimestamp As Date

'  On Error GoTo catchError
  If logFile = "" Then
    Debug.Print "Ot@CªÝè³êÄ¢Ü¹ñ"
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
    .WriteText runTime & vbTab & message, 1
    .SaveToFile logFile, 2
    .Close
  End With
  Exit Function
'G[­¶------------------------------------
catchError:
  Debug.Print "[" & Err.Number & "] OoÍ¸sF" & Err.Description
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
  
'G[­¶------------------------------------
catchError:
  Debug.Print "[" & Err.Number & "] t@CoÍ¸sF" & Err.Description
  Debug.Print "[" & Err.Number & "] " & outputFilePath
  Debug.Print "[" & Err.Number & "] " & message
End Function




'**************************************************************************************************
' * CSV`®t@CC|[g[csv/txt]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @link   https://www.tipsfound.com/vba/18014
'**************************************************************************************************
'==================================================================================================
Function importCsv(filePath As String, Optional encode As String = "sjis", Optional readLine As Long, Optional TextFormat As Variant)
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
    If encode = "sjis" Then
      .TextFilePlatform = 932
    ElseIf encode = "utf-8" Then
      .TextFilePlatform = 65001
    End If
    .TextFileParseType = xlDelimited ' ¶ÅæØÁ½`®
    .TextFileCommaDelimiter = True   ' æØè¶ÍJ}
    .TextFileStartRow = readLine     ' 1sÚ©çÇÝÝ
    .AdjustColumnWidth = False       ' ñð©®²®µÈ¢
    .RefreshStyle = xlOverwriteCells 'ã«ðwè
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' øpÌwè

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
' * t@CC|[g
Function importText(filePath As String, Optional encode As String = "sjis")
  Dim buf As String, tmp As Variant, tmpJ As Variant, i As Long, j As Long

  With CreateObject("ADODB.Stream")
    .Charset = encode
    .Open
    .LoadFromFile filePath
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

  ActiveWorkbook.Sheets(targeSheet).Range(targeArea).copy
  dictSheet.Range("A1").PasteSpecial xlPasteValues

  Application.CutCopyMode = False
  ActiveWorkbook.Close SaveChanges:=False
  dictSheet.Range("A1").Select

  DoEvents
  Call Library.startScript

    Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * pX[h¶¬
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
' * nCCg»
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setHighLight(SetArea As String, DisType As Boolean, SetColor As String)
  Range(SetArea).Select

  'ðt«®ðNA
  Selection.FormatConditions.delete

  If DisType = False Then
    's¾¯Ýè
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    'sÆñÉÝè
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
' * ¶ñª
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
' * zñÌÅãÉÇÁ·é
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
' * tHgJ[Ýè
' *
' * @Link https://vbabeginner.net/vbaÅZÌwè¶ñÌFâ¾³ðÏX·é/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font
  Dim i, iLen
  Dim R   As Range

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
' * WXgÖA
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setRegistry(RegistrySubKey As String, RegistryKey As String, setVal As Variant)
  Const funcName As String = "Library.setRegistry"
  
  'Jn--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Call Library.showDebugForm("MainKey", thisAppName, "debug")
  Call Library.showDebugForm("SubKey ", RegistrySubKey, "debug")
  Call Library.showDebugForm("Key    ", RegistryKey, "debug")
  Call Library.showDebugForm("Val    ", CStr(setVal), "debug")
  
  Call SaveSetting(thisAppName, RegistrySubKey, RegistryKey, setVal)
  
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function getRegistry(RegistryKey As String, RegistrySubKey As String, Optional typeVal As String = "String")
  Dim regVal As String
  Const funcName As String = "Library.getRegistry"

  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  
  If RegistryKey <> "" Then
    regVal = GetSetting(thisAppName, RegistryKey, RegistrySubKey)
  End If
  
  Call Library.showDebugForm("MainKey", thisAppName, "debug")
  Call Library.showDebugForm("Key    ", RegistryKey, "debug")
  Call Library.showDebugForm("SubKey ", RegistrySubKey, "debug")
  Call Library.showDebugForm("Val    ", regVal, "debug")
  Call Library.showDebugForm("type   ", typeVal, "debug")
  
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
  
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function

'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function delRegistry(RegistryKey As String, RegistrySubKey As String)
  Dim regVal As String

  Const funcName As String = "Library.delRegistry"
  On Error GoTo catchError
  'Call Library.showDebugForm(funcName, , "function1")

  If RegistrySubKey = "" Then
    Call DeleteSetting(thisAppName, RegistryKey)
  Else
    Call DeleteSetting(thisAppName, RegistryKey, RegistrySubKey)
  End If
  Exit Function

'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * V[gÌÛì/Ûìð
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
'G[­¶------------------------------------
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
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * ÅÌV[gðIð
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
' * lÌÝè
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setValandRange(keyName As String, val As String)
  Const funcName As String = "Library.setValandRange"

'  Range(keyName) = val
  If LadexsetVal Is Nothing Then
    Call init.setting
  Else
    LadexsetVal(keyName) = val
  End If
  Call Library.showDebugForm(funcName, keyName & "/" & val, "info")
End Function

'**************************************************************************************************
' * ob`t@CÀs
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function runBat(fileName As String)
  Dim obj As WshShell
  Dim rtnVal As String

  Set obj = New WshShell
  rtnVal = obj.run(fileName, WaitOnReturn:=True)

  Call Library.showDebugForm("Àst@C", fileName, "info")
  Call Library.showDebugForm("ßèl      ", rtnVal, "info")

  runBat = rtnVal
End Function

'**************************************************************************************************
' * t@CSÌÌ¶ñu·
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(fileName As String, TargetText As String, Optional NewText As String = "")
 Dim FSO         As FileSystemObject 't@CVXeIuWFNg
 Dim Txt         As TextStream       'eLXgXg[IuWFNg
 Dim buf_strTxt  As String           'ÇÝÝobt@

  Const funcName As String = "Library.replaceFromFile"
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")

 'IuWFNgì¬
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(fileName, ForReading)

 'S¶ÇÝÝ
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '³t@Cðl[µÄAe|t@Cì¬
  Name fileName As fileName & "_"

  'u·
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  'ÝpeLXgt@Cì¬
   Set Txt = FSO.CreateTextFile(fileName, True)
  'Ý
  Txt.Write buf_strTxt
  Txt.Close

  'e|t@Cðí
  FSO.DeleteFile fileName & "_"
  Set Txt = Nothing
  Set FSO = Nothing
  Exit Function

'G[­¶------------------------------------
catchError:
  FSO.DeleteFile fileName & "_"
  Set Txt = Nothing
  Set FSO = Nothing
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'**************************************************************************************************
' * VBAÅExcelÌRgðêÅ©®TCYÉµÄJbRæ­·é
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
        'TCYÝè
        .TextFrame.AutoSize = True
        .TextFrame.Characters.Font.Size = FontSizeVal
        .TextFrame.Characters.Font.Color = FontColorVal

        '`óðpÛlp`ÉÏX
        .AutoShapeType = msoShapeRectangle

        'F
        .line.ForeColor.RGB = RGB(128, 128, 128)
        .Fill.ForeColor.RGB = BgColorVal

        'e §ß¦ 30%AItZbgÊ x:1px,y:1px
        .Shadow.Transparency = 0.3
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 1

        ' ¾ðAµ¦
        .TextFrame.Characters.Font.Bold = False
        .TextFrame.HorizontalAlignment = xlLeft

        .TextFrame.Characters.Font.Name = FontVal

        ' ZÉí¹ÄÚ®·é
        .Placement = xlMove
      End With
    End If
  Next cl
End Function

'**************************************************************************************************
' * Nð
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLink()
  Dim Links As Variant
  Dim i As Integer

  Links = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks) 'ubNÌÉ éN

  If IsArray(Links) Then
    For i = 1 To UBound(Links)
      ActiveWorkbook.BreakLink Links(i), xlLinkTypeExcelLinks 'Nð
    Next i
  End If
End Function

'**************************************************************************************************
' * X[v
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  DoEvents
  Application.Wait [Now()] + timeVal / 86400000
  DoEvents
End Function

'**************************************************************************************************
' * rü
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function rü_NA(Optional SetArea As Range)
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
Function rü_\(Optional SetArea As Range, Optional LineColor As Variant)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_ÍÝ(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_iq(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_¶(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_¶E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_ã(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_º(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_ãº(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_¼(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_jü_½(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_ÍÝ(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_iq(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_¶(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
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
Function rü_Àü_E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_¶E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_ã(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_º(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_ãº(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_¼(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_Àü_½(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_ÍÝ(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_¶(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_¶E(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_ã(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_º(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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
Function rü_ñdü_ãº(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
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

'==================================================================================================
Function rü_jü_tL(Optional SetArea As Range, Optional LineColor As Variant, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  If IsMissing(LineColor) = True Then
    LineColor = LadexsetVal("LineColor")
  End If
  Call Library.getRGB(CLng(LineColor), Red, Green, Blue)
  
  Call rü_jü_ÍÝ(SetArea, LineColor, WeightVal)

  If TypeName(SetArea) = "Range" Then
    Set SetArea = SetArea.Offset(1, 1).Resize(SetArea.Rows.count - 1, SetArea.Columns.count - 1)
    Call rü_jü_½(SetArea, LineColor, WeightVal)
    Call rü_jü_ÍÝ(SetArea, LineColor, WeightVal)
  Else
    SetArea.Offset(1, 1).Resize(SetArea.Rows.count - 1, SetArea.Columns.count - 1).Select
    Call rü_jü_½(SetArea, LineColor, WeightVal)
    Call rü_jü_ÍÝ(SetArea, LineColor, WeightVal)

  End If
End Function

'==================================================================================================
Function rü_üí_¡(Optional SetArea As Range)
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
Function rü_üí_c(Optional SetArea As Range)
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
' * JÝè / æ¾
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
  
  'Jn--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "function1")
  '----------------------------------------------
  endColLine = Cells(1, Columns.count).End(xlToLeft).Column

  For colLine = 1 To endColLine
    If IsNumeric(Cells(1, colLine)) Then
      Cells(1, colLine).ColumnWidth = Cells(1, colLine)
      End If
  Next
  Exit Function
'G[­¶------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * y[WÌXe[^XmF
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

    If LadexsetVal("proxyURL") <> "" Then
      .SetProxy 2, LadexsetVal("proxyURL") & ":" & LadexsetVal("proxyPort")
    End If
    If LadexsetVal("proxyUser") <> "" Then
      .setProxyCredentials LadexsetVal("proxyUser"), LadexsetVal("proxyPasswd")
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
'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & Err.Description & ">", True)
  getURLStatusCode = 404
  Set Http = Nothing
End Function


'**************************************************************************************************
' * V[gÛì/ð
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
      If sheetName.Name Like "s*t" Then
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

  'Jn--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------

  Call Library.showDebugForm("allSheetflg", allSheetflg, "debug")
  If allSheetflg = True Then
    For Each sheetName In Sheets
      If sheetName.Name Like "s*t" Then
      Else
        ThisWorkbook.Worksheets(sheetName.Name).Unprotect passWord:=thisAppPasswd
      End If
      DoEvents
    Next
  Else
    ActiveSheet.Unprotect passWord:=thisAppPasswd
  End If

  'I¹--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & Err.Description & ">", True)
  Call Library.errorHandle
End Function


'**************************************************************************************************
' * V[gÌ\¦/ñ\¦
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
      If sheetName.Name Like "s*t" Then
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
      If sheetName.Name Like "s*t" Then
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
' * ¶ñA
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
' * 2³zñÌ1³ÚðRedim Preserve·é
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
'Rg[plÌzC[Êæ¾
Function getScrollRow()
  Dim scrollVal As Long
  Const GetWheelScrollLines = 104
  Const funcName As String = "Library.getScrollRow"

  'Jn--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

  SystemParametersInfo GetWheelScrollLines, 0, scrollVal, 0
  Call Library.setValandRange("scrollRowCnt", CStr(scrollVal))

  'I¹--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'Rg[plÌzC[ÊÝè
Function setScroll(setScrollRow As Long)
  Const SENDCHANGE = 3
  Const SetWheelScrollLines = 105
  Const funcName As String = "Library.setScroll"

  'Jn--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.showDebugForm("setScrollRow", setScrollRow, "debug")
  '----------------------------------------------

  SystemParametersInfo SetWheelScrollLines, setScrollRow, 0, SENDCHANGE


  'I¹--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * init@CÇÝÝ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getConfigIni(filePath As String)
  Dim buf As String
  Dim SectionVal As String, keyName As String, keyVal As String
  Dim keys As Variant

  Const funcName As String = "Library.getConfig"

  'Jn--------------------------------------
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
    .LoadFromFile filePath
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

  'I¹--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function setConfigIni(sectionName As String, keyName As String, filePath As String, setVal As String)
  Const funcName As String = "Library.setLineHeight"


  'Jn--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------

'  Call WritePrivateProfileString(sectionName, keyName, setVal, filePath)



  'I¹--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function getLineHeight(targetRang As Range, maxLen As Long, defaultRowHeight As Long)
  Dim LFCount As Long, LenCount As Long
'  Dim setHeight As Long
  Const funcName As String = "Library.getLineHeight"

  'Jn--------------------------------------
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

  'I¹--------------------------------------
  Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'G[­¶------------------------------------
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

  'Jn--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "function1")
  '----------------------------------------------

  Cells.EntireColumn.AutoFit

  For colLine = 1 To Columns.count
    If Cells(1, colLine).ColumnWidth > 30 Then
      colName = Library.getColumnName(colLine)
      Columns(colName & ":" & colName).ColumnWidth = 30
    End If
  Next

  'I¹--------------------------------------

  '----------------------------------------------

  Exit Function
'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function





'==================================================================================================
' * ¶ñÌ¶¤©çwè¶o
Function getLeftString(targetStr, strCnt As Long) As String
  Dim targetLen As Long
  Dim getStr As String

  targetLen = Len(targetStr)

  '¶ñÅÍÈ¢ê
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
' * ¶ñÌE¤©çwè¶o
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


