Attribute VB_Name = "Ctl_Base"
Option Explicit


'==================================================================================================
Function ChromeN®(Optional showType As String)

  Dim shellProcID  As Long
  Const funcName As String = "Ctl_Base.ChromeN®"
  
  'chromeDriver ÌXVmF-------------------------------------------------------------------------
  
  shellProcID = Shell(binPath & "\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call Library.showDebugForm("targetURL", targetURL, "debug")
  
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfilesDir)
    .AddArgument ("--window-size=1200,1200")
    
    'g£@\ð³ø
    .AddArgument ("--disable-extensions")
    
    
    If setVal("ProxyFlg") = "TRUE" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
      
      If setVal("ProxyID") <> "" Then
        .AddArgument ("--proxy-auth=" & setVal("ProxyID") & ":" & setVal("ProxyPW"))
      End If
    End If
    
    
    'SSLFØÌ³ø»
    .AddArgument ("--ignore-certificate-errors")
    
    
    Select Case showType
      Case "no-img"
        'æñ\¦
        .AddArgument ("--blink-settings=imagesEnabled=false")
      Case "app"
        .AddArgument ("--app=" & Ctl_Base.î{FØ(targetURL))
      
      Case "headless"
        .AddArgument ("--disable-gpu")
        .AddArgument ("--headless")
      
      Case Else
        'V[Nbg[h
        .AddArgument ("--incognito")
    End Select
    
    .start "Chrome"
    
    'æÊÊu
    '.Window.SetPosition Windows(targetBook.Name).Left, Windows(targetBook.Name).Top
    .Wait 1000
    
    'y[WÌ[hÌÒ¿Ô
    .Timeouts.PageLoad = 60000
    
    'javascriptÀs®¹ÌÒ¿Ô
    .Timeouts.Script = 10000
    
    'vfª©Â©éÜÅÌÒ¿Ô
    .Timeouts.ImplicitWait = 1000
    
    If targetURL <> "" Then
      .Get Ctl_Base.î{FØ(targetURL)
    End If
    
  End With
  
    
End Function

'==================================================================================================
Function ChromeI¹()
  
  Call Library.execDeldir(BrowserProfilesDir)
  driver.Quit
  Set driver = Nothing
  
  Call Library.execMkdir(BrowserProfilesDir)
  
End Function


'==================================================================================================
Function î{FØ(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.î{FØ"

  'Jn--------------------------------------
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
  
  
  'I¹--------------------------------------
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
  î{FØ = Ctl_Base.ftHgy[WÝè(baseURL)
  
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ftHgy[WÝè(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.ftHgy[WÝè"
  
  'Jn--------------------------------------
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
    
    '/ÅIíÁÄ¢éêÍAdefaultPageð·é
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
  
  'I¹--------------------------------------
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
  ftHgy[WÝè = baseURL
  
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function üÍ¶ì¬(ByVal dataVal As Variant)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim inputVal As String
  Dim DataType As Variant, DataLength As Variant, DataFormat As Variant
  
  Const funcName As String = "Ctl_Base.üÍ¶ì¬"

  'Jn--------------------------------------
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
'_Öp
'Public Const HalfWidthDigit = "1234567890"
'Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"
'
'Public Const JapaneseCharacters = " ¢¤¦¨©«­¯±³µ·¹»½¿ÂÄÆÈÉÊËÌÍÐÓÖÙÜÝÞßàâäæçèéêëíðñª¬®°²´¶¸º¼¾ÀÃÅÇÎÑÔ×ÚÏÒÕØÛ"
'Public Const JapaneseCharactersCommonUse = "JwóàÂÑæâom·¼Xà¾åéÏçjÝêïKngÀÒÌæóègÂ\¨ú¡½ûapÊè¯G¦aÅüQ¡ü¼²êITÅO{@qáÕ¿ÍøÈ}Éq³µ«§ÅìñgÄá@SÍûqÀóËïFµü¨hÀª©ñïËtÞ|øbÎ¢£"
'Public Const MachineDependentCharacters = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]_ú\ú]ú^ú_ú`úaúbúcúdúeúfúgúhúiújúkúlúmúnúoúpúqúrúsútúuúvúwúxúyúzú{"
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
    
    
    Case "URL`®"
      inputVal = LCase(HalfWidthCharacters)
      inputVal = Range("E1") & "/" & Library.makeRandomString(DataLength, inputVal)
    
    Case "ÐçªÈ"
    
      inputVal = JapaneseCharacters
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    Case "Sp¶"
      inputVal = JapaneseCharactersCommonUse
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    
    Case Else
  End Select



  Call Library.showDebugForm("inputVal", inputVal, "debug")

  'I¹--------------------------------------
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
  üÍ¶ì¬ = inputVal
  
  Exit Function

'G[­¶------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
