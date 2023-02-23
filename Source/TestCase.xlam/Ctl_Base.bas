Attribute VB_Name = "Ctl_Base"
Option Explicit


'==================================================================================================
Function Chrome起動(Optional showType As String)

  Dim shellProcID  As Long
  Const funcName As String = "Ctl_Base.Chrome起動"
  
  'chromeDriver の更新確認-------------------------------------------------------------------------
  
  shellProcID = Shell(binPath & "\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call Library.showDebugForm("targetURL", targetURL, "debug")
  
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfilesDir)
    .AddArgument ("--window-size=1200,1200")
    
    '拡張機能を無効
    .AddArgument ("--disable-extensions")
    
    
    If setVal("ProxyFlg") = "TRUE" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
      
      If setVal("ProxyID") <> "" Then
        .AddArgument ("--proxy-auth=" & setVal("ProxyID") & ":" & setVal("ProxyPW"))
      End If
    End If
    
    
    'SSL認証の無効化
    .AddArgument ("--ignore-certificate-errors")
    
    
    Select Case showType
      Case "no-img"
        '画像非表示
        .AddArgument ("--blink-settings=imagesEnabled=false")
      Case "app"
        .AddArgument ("--app=" & Ctl_Base.基本認証(targetURL))
      
      Case "headless"
        .AddArgument ("--disable-gpu")
        .AddArgument ("--headless")
      
      Case Else
        'シークレットモード
        .AddArgument ("--incognito")
    End Select
    
    .start "Chrome"
    
    '画面位置
    '.Window.SetPosition Windows(targetBook.Name).Left, Windows(targetBook.Name).Top
    .Wait 1000
    
    'ページのロードの待ち時間
    .Timeouts.PageLoad = 60000
    
    'javascript実行完了の待ち時間
    .Timeouts.Script = 10000
    
    '要素が見つかるまでの待ち時間
    .Timeouts.ImplicitWait = 1000
    
    If targetURL <> "" Then
      .Get Ctl_Base.基本認証(targetURL)
    End If
    
  End With
  
    
End Function

'==================================================================================================
Function Chrome終了()
  
  Call Library.execDeldir(BrowserProfilesDir)
  driver.Quit
  Set driver = Nothing
  
  Call Library.execMkdir(BrowserProfilesDir)
  
End Function


'==================================================================================================
Function 基本認証(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.基本認証"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  If InStr(baseURL, setVal("siteMapURL")) <> 0 Then
    Set reg = CreateObject("VBScript.RegExp")
    
    ID = setVal("authName")
    PW = setVal("authPassword")
    
    If setVal("authTypeBasic") = "TRUE" Then
      If (InStr(baseURL, "http://") > 0) Then
        baseURL = Replace(baseURL, "http://", "")
        baseURL = "http://" & ID & ":" & PW & "@" & baseURL
        
      ElseIf (InStr(baseURL, "https://") > 0) Then
        baseURL = Replace(baseURL, "https://", "")
        baseURL = "https://" & ID & ":" & PW & "@" & baseURL
        
      End If
    End If
  End If
  Call Library.showDebugForm("baseURL  ", baseURL, "debug")
  
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  基本認証 = Ctl_Base.デフォルトページ設定(baseURL)
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function デフォルトページ設定(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.デフォルトページ設定"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
 
  If InStr(baseURL, setVal("siteMapURL")) <> 0 Then
    Set reg = CreateObject("VBScript.RegExp")
    
    '/で終わっている場合は、defaultPageを結合する
    With reg
      .Pattern = "/$"
      .IgnoreCase = True
      .Global = True
    End With
    
    If reg.Test(baseURL) Then
      baseURL = baseURL & setVal("defaultPage")
      targetURL = targetURL & setVal("defaultPage")
    End If
    Call Library.showDebugForm("baseURL  ", baseURL, "debug")
    
    Set reg = Nothing
  End If
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
'    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  デフォルトページ設定 = baseURL
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function 入力文字作成(ByVal dataVal As Variant)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim inputVal As String
  Dim DataType As Variant, DataLength As Variant, DataFormat As Variant
  
  Const funcName As String = "Ctl_Base.入力文字作成"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
'ランダム関数用
'Public Const HalfWidthDigit = "1234567890"
'Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"
'
'Public Const JapaneseCharacters = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ"
'Public Const JapaneseCharactersCommonUse = "雨学空金青林画岩京国姉知長直店東歩妹明門夜委育泳岸苦具幸始使事実者昔取受所注定波板表服物放味命油和英果芽官季泣協径固刷参治周松卒底的典毒念府法牧例易往価河居券効妻枝舎述承招性制版肥非武沿延拡供呼刻若宗垂担宙忠届乳拝並宝枚依押奇祈拠況屈肩刺沼征姓拓抵到突杯泊拍迫彼怖抱肪茂炎欧殴"
'Public Const MachineDependentCharacters = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㍉纊褜鍈銈蓜俉炻昱棈鋹曻彅丨仡仼伀伃伹佖侒侊侚侔俍偀倢俿倞偆偰偂傔"
'

  DataType = dataVal(0)
  DataLength = dataVal(1)
  
  
  
  Select Case DataType
    Case "*"
      
      inputVal = HalfWidthDigit & HalfWidthCharacters & SymbolCharacters & JapaneseCharacters & JapaneseCharactersCommonUse & MachineDependentCharacters
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    Case "day"
      DataFormat = dataVal(2)
      If InStr(DataLength, "y") >= 1 Then
      ElseIf InStr(DataLength, "m") >= 1 Then
        DataLength = Replace(DataLength, "m", "")
        inputVal = DateAdd("m", DataLength, Now())
      Else
        inputVal = DateAdd("d", DataLength, Now())

      End If
      inputVal = Format(inputVal, DataFormat)
    
    
    Case "URL形式"
      inputVal = LCase(HalfWidthCharacters)
      inputVal = Range("E1") & "/" & Library.makeRandomString(DataLength, inputVal)
    
    Case "ひらがな"
    
      inputVal = JapaneseCharacters
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    Case "全角文字"
      inputVal = JapaneseCharactersCommonUse
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    
    Case Else
  End Select



  Call Library.showDebugForm("inputVal", inputVal, "debug")

  '処理終了--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  入力文字作成 = inputVal
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
