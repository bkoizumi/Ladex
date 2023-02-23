Attribute VB_Name = "Ctl_Selenium"
Option Explicit

Dim ActivLine As Long
Dim Target    As String


Dim Cmd       As String
Dim targetVal As String
Dim waitFlg   As Boolean
Dim TestCaseName   As String

Dim DataType    As String
Dim DataLength  As Integer
Dim DataReqFlg As Boolean

Dim resultCell As Long
Dim evidenceCell As Long


Const defPageHeight As Long = 1200
Const defPageWidth As Long = 1200


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 開始()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim getEvidenceFlg As Boolean
  Const funcName As String = "Ctl_Selenium.開始"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
'    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  Set targetBook = ActiveWorkbook
  

  endLine = Range("A1").SpecialCells(xlLastCell).Row
  
  Call Ctl_TestCase.セル範囲設定(startLine, endLine)
  
  '結果入力セルの設定----------------------------
    Range(resultArea1).ClearContents
    Range(resultArea1).ClearComments
    
    Range(resultArea2).ClearContents
    Range(resultArea2).ClearComments
    
    Range(resultArea3).ClearContents
    Range(resultArea3).ClearComments
    
    Range(resultArea4).ClearContents
    Range(resultArea4).ClearComments
    
    Range(resultArea5).ClearContents
    Range(resultArea5).ClearComments
    
    resultCell = 15
    evidenceCell = 19

'  If Range("Q9").value <> Range("Q10").value Then
'    Range(resultArea1).ClearContents
'    Range(resultArea1).ClearComments
'    resultCell = 15
'    evidenceCell = 19
'
'  ElseIf Range("W9").value <> Range("W10").value Then
'    Range(resultArea2).ClearContents
'    Range(resultArea2).ClearComments
'    resultCell = 21
'    evidenceCell = 25
'
'  ElseIf Range("AC9").value <> Range("AC10").value Then
'    Range(resultArea3).ClearContents
'    Range(resultArea3).ClearComments
'    resultCell = 27
'    evidenceCell = 31
'
'  ElseIf Range("AI9").value <> Range("AI10").value Then
'    Range(resultArea4).ClearContents
'    Range(resultArea4).ClearComments
'    resultCell = 33
'    evidenceCell = 37
'
'  ElseIf Range("AO9").value <> Range("AO10").value Then
'    Range(resultArea5).ClearContents
'    Range(resultArea5).ClearComments
'    resultCell = 39
'    evidenceCell = 43
'  End If
  
  
  
  
  Call Ctl_Base.Chrome起動
  
  For ActivLine = startLine To endLine
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ActivLine - startLine, endLine - startLine, "selenium実行中")
    getEvidenceFlg = False
    
    
    Select Case Range("H" & ActivLine).value
      Case "画面遷移"
        Call Ctl_Selenium.画面遷移
      
      Case "文字入力"
        Call Ctl_Selenium.文字入力
      
      Case "表示確認"
        Call Ctl_Selenium.表示確認
      
      Case "ファイル選択"
        Call Ctl_Selenium.ファイル選択
      
      Case "チェックボックス選択"
        Call Ctl_Selenium.チェックボックス選択
      
      Case "ラジオボタン選択"
        Call Ctl_Selenium.ラジオボタン選択
      
      
      Case "リンククリック"
        Call Ctl_Selenium.リンククリック(True)
      
      Case "スクリｰンショット"
          getEvidenceFlg = True
      
      Case "プルダウン選択"
          Call Ctl_Selenium.プルダウン選択
      
      Case "ボタンクリック"
          getEvidenceFlg = True
          Call Ctl_Selenium.ボタンクリック(True)
      
      Case "手動確認/操作"
        Call Ctl_Selenium.手動確認

      Case ""
  
      Case Else
    End Select
    
    Select Case Range("H" & ActivLine).value
      Case "", "表示確認", "手動確認/操作"
      
      
      Case Else
        Call Ctl_Selenium.画面キャプチャ(getEvidenceFlg)
    End Select

  Next
  driver.Quit
  Set driver = Nothing


'  Application.Goto Reference:=Range("A1"), Scroll:=True
  '処理終了--------------------------------------
  If runFlg = False Then
    
    Call Ctl_ProgressBar.showEnd
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
Function 手動確認()
  Dim meg As String
  
  
  Const funcName As String = "Ctl_Selenium.手動確認"

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.endScript
  '----------------------------------------------
  
  
  Application.Goto Reference:=Range("A" & ActivLine), Scroll:=True
  DoEvents
  
  meg = "確認をお願いします。"
  
  If Range("J" & ActivLine) <> "" Then
    meg = "手動操作および確認をお願いします。"
  End If
  
  If Range("H" & ActivLine - 1) <> "手動確認/操作" Then
    Application.Speech.Speak Text:=meg, SpeakAsync:=True, SpeakXML:=True
  End If
  
  With Frm_Wait
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    
'    .Top = Range("J" & ActivLine).Top
'    .Left = Range("J" & ActivLine).Top
    
    .Caption = ActivLine & "行目 " & meg
    .TextBox3.value = Range("J" & ActivLine)
    .TextBox2.value = Range("L" & ActivLine)
    .Show
  End With


  '処理終了--------------------------------------
    Call Library.startScript
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function 画面遷移()
  Const funcName As String = "Ctl_Selenium.画面遷移"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  Target = Range("I" & ActivLine)
  
  Call Library.showDebugForm("Target", Target, "debug")
  driver.Get Target
  
  If driver.title = "プライバシー エラー" Then
    driver.FindElementById("details-button").Click
    driver.FindElementById("proceed-link").Click
  End If
  Call Ctl_Selenium.テスト結果(True)
  

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function



'==================================================================================================
Function 文字入力()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim testStrMinLen As Variant, testStrMaxLen As Variant, testStrType As String
   
  Const funcName As String = "Ctl_Selenium.文字入力"

  '処理開始--------------------------------------
'  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  
  
  Call Library.showDebugForm("Target", Target, "debug")
  
  tmpVal = Split(Target, "=")
  key = tmpVal(0)
  value = tmpVal(1)
  Call Library.showDebugForm("key   ", key, "debug")
  
  
  element = Split(key, ".")
  elementType = element(0)
  elementName = element(1)
  Call Library.showDebugForm("elementType", elementType, "debug")
  Call Library.showDebugForm("elementName", elementName, "debug")
  Call Library.showDebugForm("value      ", value, "debug")
  
  
  If InStr(value, "auto-") >= 1 Then
    value = Replace(value, "auto-", "")
    tmpVal = Split(value, ",")
    
    value = Ctl_Base.入力文字作成(tmpVal)
    Cells(ActivLine, resultCell + 3) = "入力値：" & value
  End If
  
  Select Case LCase(elementType)
    Case "name"
      driver.FindElementByName(elementName).Clear
      driver.FindElementByName(elementName).SendKeys value
    Case "class"
      driver.FindElementByClass(elementName).Clear
      driver.FindElementByClass(elementName).SendKeys value
    Case "id"
      driver.FindElementById(elementName).Clear
      driver.FindElementById(elementName).SendKeys value
  End Select
  Call Ctl_Selenium.テスト結果(True)
  

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function



'==================================================================================================
Function 表示確認()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim chkFlg As Boolean
  Dim getTxt As String
  Dim elements  As WebElements
  
  
  Const funcName As String = "Ctl_Selenium.表示確認"

  '処理開始--------------------------------------
'  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  If InStr(Target, "=") >= 1 Then
    tmpVal = Split(Target, "=")
    key = tmpVal(0)
    value = tmpVal(1)
    Call Library.showDebugForm("key   ", key, "debug")
  Else
    key = Target
  End If
  
  If InStr(key, ".") >= 1 Then
    element = Split(key, ".")
    elementType = element(0)
    elementName = element(1)
    Call Library.showDebugForm("elementType", elementType, "debug")
    Call Library.showDebugForm("elementName", elementName, "debug")
    Call Library.showDebugForm("value      ", value, "debug")
  Else
    elementType = element(0)
    elementName = element(1)
  End If
    
  Set elements = Nothing
  Select Case LCase(elementType)
    Case "name"
      getTxt = driver.FindElementByName(elementName).Text()
    Case "class"
      If driver.FindElementsByClass(elementName).count = 1 Then
        getTxt = driver.FindElementByClass(elementName).Text()
      Else
        Set elements = driver.FindElementsByClass(elementName)
      End If
    Case "id"
      getTxt = driver.FindElementById(elementName).Text()
    Case Else
  End Select
  
  If InStr(value, "auto-") >= 1 Then
    value = Replace(value, "auto-", "")
    tmpVal = Split(value, ",")
    value = "*" & Ctl_Base.入力文字作成(tmpVal) & "*"
  End If
  
  Call Library.showDebugForm("getTxt     ", getTxt, "debug")
  
  If elements Is Nothing Then
    If InStr(value, "*") >= 1 Then
      If getTxt Like value Then
        Call Ctl_Selenium.テスト結果(True)
      Else
         Call Ctl_Selenium.テスト結果(False)
      End If
    Else
      If getTxt = value Then
        Call Ctl_Selenium.テスト結果(True)
      Else
         Call Ctl_Selenium.テスト結果(False)
      End If
    End If
  Else
    For Each element In elements
      Call Library.showDebugForm("getTxt     ", element.Text, "debug")
      
      If InStr(value, "*") >= 1 Then
        If element.Text Like value & "*" Then
          Call Ctl_Selenium.テスト結果(True)
          Exit For
        End If
      Else
        If element.Text = value Then
          Call Ctl_Selenium.テスト結果(True)
          Exit For
        End If
      End If
    Next
  End If
  
  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function ボタンクリック(waitFlg As Boolean)
  Dim cmdVal As Variant
  Dim key As String, value As String
  Dim myBy As New BY
  Dim element As Variant
  Dim elementType As String, elementName As String
  
  Const funcName As String = "Ctl_Selenium.ボタンクリック"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  element = Split(Target, ".")
  elementType = element(0)
  elementName = element(1)
  Call Library.showDebugForm("elementType", elementType, "debug")
  Call Library.showDebugForm("elementName", elementName, "debug")
  
  driver.FindElementByTag("body").SendKeys vbTab
  Select Case LCase(elementType)
    Case "name"
      driver.FindElementByName(elementName).Click
    
    Case "class"
      driver.FindElementByClass(elementName).Click
    
    Case "id"
      driver.FindElementById(elementName).Click
    
    Case "xpath"
      driver.FindElementByXPath(elementName).Click
  
    Case Else
  End Select
  
  If waitFlg = True Then
    Do Until waitFlg = True
      waitFlg = driver.IsElementPresent(myBy.XPath("/html"))
      driver.Wait 1000
    Loop
  End If
  Call Ctl_Selenium.テスト結果(True)
  

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Ctl_Selenium.テスト結果(False)
End Function


'==================================================================================================
Function 画面キャプチャ(getEvidenceFlg As Boolean)
  Dim imgSavePath As String, imgSaveName As String
  Dim pageWidth As Long, pageHeight As Long

  
  Const funcName As String = "Ctl_Selenium.画面キャプチャ"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  TestCaseName = Range("D9") & "_" & ActivLine
  
  Call Library.showDebugForm("Target      ", Target, "debug")
  Call Library.showDebugForm("TestCaseName", TestCaseName, "debug")
  
  pageWidth = driver.ExecuteScript("return document.body.scrollWidth") + 10
  pageHeight = driver.ExecuteScript("return document.body.scrollHeight") + 10
  
  If pageWidth < defPageWidth Then pageWidth = 1200
  If pageHeight < defPageHeight Then pageHeight = 1200
  
  driver.Window.SetSize pageWidth, pageHeight
  
  
  imgSavePath = ActiveWorkbook.path & "\エビデンス"
  imgSaveName = TestCaseName & ".png"
  
  Call Library.execMkdir(imgSavePath)
  Call Library.showDebugForm("imgSaveName", imgSavePath & "\" & imgSaveName, "debug")
  
  driver.TakeScreenshot.SaveAs imgSavePath & "\" & imgSaveName

  driver.Window.SetSize defPageWidth, defPageHeight

  'エビデンの設定
  If getEvidenceFlg = True Then
    With Cells(ActivLine, resultCell + 4)
      If TypeName(.Comment) = "Comment" Then
        .ClearComments
      End If
      
      .value = imgSaveName
      With .AddComment
        .Shape.Fill.UserPicture imgSavePath & "\" & imgSaveName
        .Shape.Height = 500
        .Shape.Width = 500
      End With
    End With
  End If

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Stop
End Function


'==================================================================================================
Function テスト結果(resultFlg As Boolean, Optional strMeg As String)
  Dim line As Long
  
  Const funcName As String = "Ctl_Selenium.テスト結果"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  If resultFlg = True Then
     Cells(ActivLine, resultCell) = "OK"
  Else
    Cells(ActivLine, resultCell) = "NG"
  End If
  
  Cells(ActivLine, resultCell + 1) = Format(Date, "yyyy/mm/dd")
  Cells(ActivLine, resultCell + 2) = Application.UserName
  
  If strMeg <> "" Then
    Cells(ActivLine, resultCell + 3) = strMeg
  End If

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  If resultFlg = False Then
    Application.Goto Reference:=Range("G" & ActivLine), Scroll:=True
    Application.Speech.Speak Text:="テスト結果NG", SpeakAsync:=True, SpeakXML:=True
    Stop
  End If
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Stop
End Function


'==================================================================================================
Function ファイル選択()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim chkFlg As Boolean
  Dim getTxt As String
  Dim elements  As WebElements
  
  
  Const funcName As String = "Ctl_Selenium.表示確認"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  If InStr(Target, "=") >= 1 Then
    tmpVal = Split(Target, "=")
    key = tmpVal(0)
    value = tmpVal(1)
    Call Library.showDebugForm("key   ", key, "debug")
  Else
    key = Target
  End If
  
  If InStr(key, ".") >= 1 Then
    element = Split(key, ".")
    elementType = element(0)
    elementName = element(1)
    Call Library.showDebugForm("elementType", elementType, "debug")
    Call Library.showDebugForm("elementName", elementName, "debug")
    Call Library.showDebugForm("value      ", value, "debug")
  Else
    elementType = element(0)
    elementName = element(1)
  
  End If
    
  Select Case LCase(elementType)
    Case "name"
      driver.FindElementByName(elementName).SendKeys value
    Case "class"
      driver.FindElementByClass(elementName).SendKeys value
    Case "id"
      driver.FindElementById(elementName).SendKeys value
  End Select
  Call Ctl_Selenium.テスト結果(True)
  driver.Wait 1000

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function ラジオボタン選択()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim chkFlg As Boolean
  Dim getTxt As String
  Dim elements  As WebElements
  
  
  Const funcName As String = "Ctl_Selenium.ラジオボタン選択"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  If InStr(Target, "=") >= 1 Then
    tmpVal = Split(Target, "=")
    key = tmpVal(0)
    value = tmpVal(1)
    Call Library.showDebugForm("key   ", key, "debug")
  Else
    key = Target
  End If
  
  If InStr(key, ".") >= 1 Then
    element = Split(key, ".")
    elementType = element(0)
    elementName = element(1)
    Call Library.showDebugForm("elementType", elementType, "debug")
    Call Library.showDebugForm("elementName", elementName, "debug")
    Call Library.showDebugForm("value      ", value, "debug")
  Else
    elementType = element(0)
    elementName = element(1)
  
  End If
    
  Select Case LCase(elementType)
    Case "name"
      Set elements = driver.FindElementsByName(elementName)
    Case "class"
      Set elements = driver.FindElementsByClass(elementName)
    Case "id"
      Set elements = driver.FindElementsById(elementName)
  End Select
  
  For Each element In elements
    If element.value = value Then
      element.Click
      Exit For
    End If
  Next
  
  
  Call Ctl_Selenium.テスト結果(True)
  driver.Wait 1000

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function



'==================================================================================================
Function チェックボックス選択()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim chkFlg As Boolean
  Dim getTxt As String, chkVal As Variant
  Dim elements  As WebElements
  Dim chkMeg As String
  
  
  Const funcName As String = "Ctl_Selenium.チェックボックス選択"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  If InStr(Target, "=") >= 1 Then
    tmpVal = Split(Target, "=")
    key = tmpVal(0)
    value = tmpVal(1)
    Call Library.showDebugForm("key   ", key, "debug")
  Else
    key = Target
  End If
  
  If InStr(key, ".") >= 1 Then
    element = Split(key, ".")
    elementType = element(0)
    elementName = element(1)
    Call Library.showDebugForm("elementType", elementType, "debug")
    Call Library.showDebugForm("elementName", elementName, "debug")
    Call Library.showDebugForm("value      ", value, "debug")
  Else
    elementType = element(0)
    elementName = element(1)
  
  End If
    
  Set elements = Nothing
  Select Case LCase(elementType)
    Case "name"
      Set elements = driver.FindElementdByName(elementName)
    Case "class"
      Set elements = driver.FindElementsByClass(elementName)
    Case "id"
      Set elements = driver.FindElementsById(elementName)
  Case Else
  End Select
  
  Call Library.showDebugForm("getTxt     ", getTxt, "debug")
  
  chkFlg = False
  chkMeg = ""
  
  Select Case LCase(elementType)
    Case "name"
      
    Case "class"
      
    Case "id"
      For Each element In elements
        If element.IsSelected = False And value = 1 Then
          element.Click
          chkFlg = True
        ElseIf element.IsSelected = True And value = 0 Then
          element.Click
          chkFlg = True
        
        ElseIf element.IsSelected = True And value = 1 Then
          chkFlg = True
          chkMeg = "すでに選択されている状態"
        End If
      Next
  Case Else
  End Select
  
  Call Ctl_Selenium.テスト結果(chkFlg, chkMeg)
    
'  driver.Wait 1000

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function リンククリック(waitFlg As Boolean)
  Dim cmdVal As Variant
  Dim key As String, value As String
  Dim myBy As New BY
  Dim element As Variant
  Dim elementType As String, elementName As String
  
  Const funcName As String = "Ctl_Selenium.リンククリック"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  driver.FindElementByLinkText(Target).Click

  
  If waitFlg = True Then
    Do Until waitFlg = True
      waitFlg = driver.IsElementPresent(myBy.XPath("/html"))
      driver.Wait 1000
    Loop
  End If
  Call Ctl_Selenium.テスト結果(True)
  

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Ctl_Selenium.テスト結果(False)
End Function



'==================================================================================================
Function プルダウン選択()
  Dim tmpVal As Variant
  Dim key As String, value As String
  Dim element As Variant
  Dim elementType As String, elementName As String
  Dim chkFlg As Boolean
  Dim getTxt As String, chkVal As Variant
  Dim elements  As WebElements
  
  
  
  Const funcName As String = "Ctl_Selenium.プルダウン選択"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Target = Range("I" & ActivLine)
  Call Library.showDebugForm("Target", Target, "debug")
  
  If InStr(Target, "=") >= 1 Then
    tmpVal = Split(Target, "=")
    key = tmpVal(0)
    value = tmpVal(1)
    Call Library.showDebugForm("key   ", key, "debug")
  Else
    key = Target
  End If
  
  If InStr(key, ".") >= 1 Then
    element = Split(key, ".")
    elementType = element(0)
    elementName = element(1)
    Call Library.showDebugForm("elementType", elementType, "debug")
    Call Library.showDebugForm("elementName", elementName, "debug")
    Call Library.showDebugForm("value      ", value, "debug")
  Else
    elementType = element(0)
    elementName = element(1)
  
  End If
    
  Select Case LCase(elementType)
    Case "name"
      driver.FindElementdByName(elementName)(1).AsSelect.SelectByText (value)
    Case "class"
      driver.FindElementsByClass(elementName)(1).AsSelect.SelectByText (value)
    Case "id"
      driver.FindElementsById(elementName)(1).AsSelect.SelectByText (value)
  Case Else
  End Select
  
  Call Ctl_Selenium.テスト結果(True)

  '処理終了--------------------------------------
    Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Ctl_Selenium.テスト結果(False)
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

