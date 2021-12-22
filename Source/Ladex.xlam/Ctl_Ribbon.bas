Attribute VB_Name = "Ctl_Ribbon"
Option Explicit

Private Ctl_Event As New Ctl_Event

#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If
'**************************************************************************************************
' * リボンメニュー初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'読み込み時処理
Function onLoad(ribbon As IRibbonUI)
  Const funcName As String = "Ctl_Ribbon.onLoad"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "start")
  '----------------------------------------------
  
  Set BK_ribbonUI = ribbon
  
  BKh_rbPressed = Library.getRegistry("Main", "HighLightFlg", "Boolean")
  BKz_rbPressed = Library.getRegistry("Main", "ZoomFlg", "Boolean")
  BKT_rbPressed = Library.getRegistry("Main", "CustomRibbon", "Boolean")
  
  Call Library.showDebugForm("BKh_rbPressed", CStr(BKh_rbPressed), "debug")
  Call Library.showDebugForm("BKz_rbPressed", CStr(BKz_rbPressed), "debug")
  Call Library.showDebugForm("BKT_rbPressed", CStr(BKT_rbPressed), "debug")
  
  Call Library.setRegistry("Main", "BK_ribbonUI", CStr(ObjPtr(BK_ribbonUI)))
  Call Main.InitializeBook
  
  'BK_ribbonUI.ActivateTab ("Ladex")
  BK_ribbonUI.Invalidate
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

''==================================================================================================
'Function dMenuRefresh(control As IRibbonControl)
'  If BK_ribbonUI Is Nothing Then
'    #If VBA7 And Win64 Then
'      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
'    #Else
'      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
'    #End If
'  End If
'  BK_ribbonUI.Invalidate
'End Function

'==================================================================================================
Function getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
  Call RefreshRibbon
End Function

'==================================================================================================
Function RefreshRibbon()
  Const funcName As String = "Ctl_Ribbon.RefreshRibbon"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  'Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  BK_ribbonUI.Invalidate

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description & ">", "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If

  Dim ribbonObj As Object
  
  If lRibbonPointer = 0 Then
    End
  End If
  
'  Stop
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p)
End Function

'**************************************************************************************************
' * トグルボタンにチェックを設定する
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'ハイライト
'==================================================================================================
Function HighLight(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.HighLight"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
  BKh_rbPressed = pressed
  Call Library.setRegistry("Main", "HighLightFlg", pressed)
  
  Call Ctl_HighLight.showStart(ActiveCell)
  If pressed = False Then
    Call Library.delRegistry("Main", "HighLightFlg")
  End If
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function HighLightPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKh_rbPressed
End Function

' ズーム
'==================================================================================================
Function Zoom(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.Zoom"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
  BKz_rbPressed = pressed
  If pressed = False Then
    Call Application.OnKey("{F2}")
    Call Library.delRegistry("Main", "ZoomFlg")
  Else
    Call Application.OnKey("{F2}", "Ctl_Zoom.ZoomIn")
  End If
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ZoomInPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKz_rbPressed
End Function

'==================================================================================================
' 計算式確認
Function confirmFormula(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.confirmFormula"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  BKcf_rbPressed = pressed
  
  Call Ctl_Formula.数式確認
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function confFormulaPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKcf_rbPressed
End Function


'==================================================================================================
'お気に入りファイルを開く
Function FavoriteFileOpen(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  Dim objFso As New FileSystemObject
  Const funcName As String = "Ctl_Ribbon.FavoriteFileOpen"
  
  fileNamePath = Library.getRegistry("FavoriteList", control.ID)
  
  If Library.chkFileExists(fileNamePath) Then
    Select Case objFso.GetExtensionName(fileNamePath)
      Case "xls", "xlsx", "xlsm"
        Workbooks.Open fileName:=fileNamePath
      Case Else
        CreateObject("Shell.Application").ShellExecute fileNamePath
      End Select
  Else
    Call Library.showNotice(404, fileNamePath, True)
  End If
End Function

'**************************************************************************************************
' * リボンメニュー表示/非表示切り替え
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function noDispTab(control As IRibbonControl)
  Call Library.setRegistry("Main", "CustomRibbon", False)
  Call RefreshRibbon
End Function

'==================================================================================================
Function setDispTab(control As IRibbonControl, pressed As Boolean)
  BKT_rbPressed = pressed
  Call Library.setRegistry("Main", "CustomRibbon", pressed)
  Call RefreshRibbon
End Function

'==================================================================================================
Function getDispTab(control As IRibbonControl, ByRef returnedVal)
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
End Function

'**************************************************************************************************
' * ダイナミックメニュー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'シート一覧メニュー
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim SheetName As Worksheet
  Dim MenuSepa, sheetNameID
  
'  On Error GoTo catchError
  If Workbooks.count = 0 Then
    Call MsgBox("ブックが開かれていません", vbCritical, thisAppName)
    Call Library.endScript
    End
  End If
  Call init.setting
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "シート管理"
      .SetAttribute "title", "シート管理"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "シート管理表示"
      .SetAttribute "label", "シート管理"
      .SetAttribute "supertip", "シート管理"
      
      .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .SetAttribute "onAction", "Ladex.xlam!Menu.ladex_シート管理_フォーム表示"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
      With MenuSepa
        .SetAttribute "id", "M_RelaxTools"
        .SetAttribute "title", "RelaxToolsを利用"
      End With
      Menu.AppendChild MenuSepa
      Set MenuSepa = Nothing

      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "RelaxTools"
        .SetAttribute "label", "RelaxTools"
        .SetAttribute "supertip", "RelaxToolsのシート管理を起動"
        
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
        .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
  End If
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "sheetID_0"
      .SetAttribute "title", ActiveWorkbook.Name
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
  
  
  
  For Each SheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = SheetName.Name
      .SetAttribute "id", "sheetID_" & SheetName.Index
      .SetAttribute "label", SheetName.Name
    
      If ActiveWorkbook.ActiveSheet.Name = SheetName.Name Then
        .SetAttribute "supertip", "アクティブシート"
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
        
      ElseIf Sheets(SheetName.Name).Visible = True Then
       '.SetAttribute "supertip", "アクティブシート"
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      
      ElseIf Sheets(SheetName.Name).Visible = 0 Then
        .SetAttribute "supertip", "非表示シート"
        .SetAttribute "imageMso", "SheetProtect"
      
      ElseIf Sheets(SheetName.Name).Visible = 2 Then
        .SetAttribute "supertip", "マクロによる非表示シート"
        .SetAttribute "imageMso", "ReviewProtectWorkbook"
      
      End If
      
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  'Debug.Print DOMDoc.XML
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
' お気に入りメニュー
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa, tmp
  Const funcName As String = "Ctl_Ribbon.FavoriteMenu"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menuの作成

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

'  Call Ctl_Favorite.getList
'  endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
  With MenuSepa
    .SetAttribute "id", "MS_お気に入り一覧"
    .SetAttribute "title", "お気に入り一覧"
  End With
  Menu.AppendChild MenuSepa
  Set MenuSepa = Nothing
  If Not IsEmpty(tmp) Then
    For line = 0 To UBound(tmp)
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", tmp(line, 0)
        .SetAttribute "label", objFso.GetFileName(tmp(line, 1))
        
        'アイコンの設定
        Select Case objFso.GetExtensionName(tmp(line, 1))
          Case "xlsm"
            .SetAttribute "imageMso", "MicrosoftExcel"
          Case "pdf"
            .SetAttribute "imageMso", "FileSaveAsPdf"
          Case "docs"
            .SetAttribute "imageMso", "FileSaveAsWordDocx"
          Case "text"
            .SetAttribute "imageMso", "FileNewContext"
          Case "accdb"
            .SetAttribute "imageMso", "MicrosoftAccess"
          Case "pptx"
            .SetAttribute "imageMso", "MicrosoftPowerPoint"
          Case "csv"
            .SetAttribute "imageMso", "FileNewContext"
          
          Case Else
            .SetAttribute "imageMso", "MicrosoftExcel"
            Call Library.showDebugForm("お気に入りアイコン", objFso.GetExtensionName(tmp(line, 1)), "Error")
        End Select
        
        
        .SetAttribute "supertip", tmp(line, 1)
        .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteFileOpen"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
    Next
  End If
  DOMDoc.AppendChild Menu
  returnedVal = DOMDoc.XML
'  Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
'RelaxTools
Function getRelaxTools(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim SheetName As Worksheet
  Dim MenuSepa

  Const funcName As String = "Ctl_Ribbon.getRelaxTools"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
    
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"
  
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    'RelaxTools取得------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxToolsGet"
      .SetAttribute "title", "RelaxToolを入手"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools_get"
      .SetAttribute "label", "RelaxToolを入手"
      .SetAttribute "image", "RelaxToolsLogo"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxTools----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxTools"
      .SetAttribute "title", "RelaxToolsを利用"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools01"
      .SetAttribute "label", "シート管理"
      .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools02"
      .SetAttribute "label", "書式リフレッシュ"
      '.SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools02"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxShapes---------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxShapes"
      .SetAttribute "title", "RelaxShapesを利用"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes01"
      .SetAttribute "label", "サイズ合わせ"
      .SetAttribute "imageMso", "ShapesDuplicate"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes01"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes02"
      .SetAttribute "label", "上位置合わせ"
      .SetAttribute "imageMso", "ObjectsAlignTop"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes02"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes03"
      .SetAttribute "label", "左位置合わせ"
      .SetAttribute "imageMso", "ObjectsAlignLeft"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes03"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxApps-----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxApps"
      .SetAttribute "title", "RelaxAppsを利用"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxApps01"
      .SetAttribute "label", "逆Ｌ罫線"
      .SetAttribute "imageMso", "BorderDrawGrid"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxApps01"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Else
    'RelaxTools取得------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxTools"
      .SetAttribute "title", "RelaxToolを入手"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools_get"
      .SetAttribute "label", "RelaxToolを入手"
      .SetAttribute "image", "RelaxToolsLogo"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  End If

  DOMDoc.AppendChild Menu
  'Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  BK_ribbonUI.InvalidateControl control.ID

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function selectActiveSheet(control As IRibbonControl)
  Dim sheetNameID As Integer
  Dim sheetCount As Integer
  Dim SheetName As Worksheet
  Const funcName As String = "Ctl_Ribbon.selectActiveSheet"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  sheetNameID = Replace(control.ID, "sheetID_", "")
  
  If Sheets(sheetNameID).Visible <> 2 Then
    Sheets(sheetNameID).Visible = True
  
  ElseIf Sheets(sheetNameID).Visible = 2 Then
    If MsgBox("マクロによって非表示となっているシートです" & vbNewLine & "マクロの動作に影響を与える可能性があります。" & vbNewLine & "表示しますか？", vbYesNo + vbCritical) = vbNo Then
      End
    Else
      Sheets(sheetNameID).Visible = True
    End If
  End If
  
  sheetCount = 1
  For Each SheetName In ActiveWorkbook.Sheets
    If Sheets(SheetName.Name).Visible = True And SheetName.Name = Sheets(sheetNameID).Name Then
      Exit For
    Else
      sheetCount = sheetCount + 1
    End If
  Next
  ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
  ActiveWindow.ScrollWorkbookTabs Sheets:=sheetCount
  Sheets(sheetNameID).Select
  
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  
'  If BK_ribbonUI Is Nothing Then
'  Else
'    BK_ribbonUI.Invalidate
'  End If
  
  Call Library.endScript
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function encode(strVal As String)
  strVal = Replace(strVal, "(", "BK1_")
  strVal = Replace(strVal, ")", "BK2_")
  strVal = Replace(strVal, " ", "BK3_")
  strVal = Replace(strVal, "　", "BK4_")
  strVal = Replace(strVal, "【", "BK5_")
  strVal = Replace(strVal, "】", "BK6_")
  strVal = Replace(strVal, "（", "BK7_")
  strVal = Replace(strVal, "）", "BK8_")
  strVal = "BK0_" & strVal
  
  encode = strVal
End Function

'==================================================================================================
Function decode(strVal As String)
  strVal = Replace(strVal, "BK0_", "")
  strVal = Replace(strVal, "BK1_", "(")
  strVal = Replace(strVal, "BK2_", ")")
  strVal = Replace(strVal, "BK3_", " ")
  strVal = Replace(strVal, "BK4_", "　")
  strVal = Replace(strVal, "BK5_", "【")
  strVal = Replace(strVal, "BK6_", "】")
  strVal = Replace(strVal, "BK7_", "（")
  strVal = Replace(strVal, "BK8_", "）")
  
  decode = strVal
End Function

'**************************************************************************************************
' * リボンメニュー[その他]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function setShrinkToFit(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.ShrinkToFit = True
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function unsetShrinkToFit(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.ShrinkToFit = False
  End If
End Function

'**************************************************************************************************
' * リボンメニュー[RelaxTools]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function RelaxTools_get(control As IRibbonControl)
  CreateObject("WScript.Shell").run ("chrome.exe -url " & "https://software.opensquare.net/relaxtools/")
End Function


'==================================================================================================
Function RelaxTools01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!execSheetManager"
End Function

'==================================================================================================
Function RelaxTools02(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!basSelection.execSelectionToFormula"
End Function

'==================================================================================================
'サイズ合わせ
Function RelaxShapes01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeSize"
End Function

'==================================================================================================
'上位置合わせ
Function RelaxShapes02(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeTop"
End Function

'==================================================================================================
'左位置合わせ
Function RelaxShapes03(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeLeft"
End Function


'==================================================================================================
'逆Ｌ罫線
Function RelaxApps01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!execSelectionFormatCheckList"
End Function

'**************************************************************************************************
' * リボンメニュー[カスタマイズ]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "start")
  
  If control.ID <> "行例入れ替えて貼付け" Then
    Call Library.startScript
  End If
  '----------------------------------------------
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "Favorite_detail"
      Call Ctl_Favorite.detail
    Case "お気に入り追加"
      Call Ctl_Favorite.add
    
    Case "Notation_R1C1"
      Call Ctl_Sheet.R1C1表記
    
    'Option--------------------------------------
    Case "Option"
      Call Ctl_Option.showOption
      
    Case "スタイル出力"
      Call Ctl_Style.Export
    Case "スタイル取込"
      Call Ctl_Style.Import
    
    Case "OptionHighLight"
      Ctl_Option.HighLight
    
    Case "OptionComment"
      Ctl_Option.Comment
    
    Case "Version"
      Call Ctl_Option.showVersion
    
    Case "Help"
      Call Ctl_Option.showHelp
    
    Case "OptionSheetImport"
      Call Ctl_RbnMaint.OptionSheetImport
    Case "OptionSheetExport"
      Call Ctl_RbnMaint.OptionSheetExport
    
    'ブック管理----------------------------------
    Case "resetStyle"
      Call Ctl_Style.スタイル初期化
    Case "delStyle"
      Call Ctl_Style.スタイル削除
    Case "setStyle"
      Call Ctl_Style.スタイル設定
    Case "del_CellNames"
      Call Ctl_Book.名前定義削除
    Case "disp_SVGA12"
      Call Ctl_Window.画面サイズ変更(612, 432)
    Case "disp_HD15_6"
      Call Ctl_Window.画面サイズ変更(1920, 1080)
    Case "シート一覧取得"
      Call Ctl_Book.シートリスト取得
    
    'シート管理----------------------------------
    Case "セル選択"
      Application.GoTo Reference:=Range("A1"), Scroll:=True
    Case "セル選択_保存"
      Application.GoTo Reference:=Range("A1"), Scroll:=True
      ActiveWorkbook.Save
    Case "全セル表示"
      Call Ctl_Sheet.すべて表示
    Case "セルとシート選択"
      Call Ctl_Sheet.A1セル選択
    Case "セルとシート_保存"
      Call Ctl_Sheet.A1セル選択
      ActiveWorkbook.Save
    Case "標準画面"
      Call Ctl_Sheet.標準画面
    Case "シート管理"
      Call Ctl_Sheet.シート管理_フォーム表示
    
    'ズーム--------------------------------------
    Case "Zoom01"
      Call Ctl_Zoom.Zoom01
    
    'セル調整------------------------------------
    Case "セル調整_幅"
      Call Ctl_Sheet.セル幅調整
    Case "セル調整_高さ"
      Call Ctl_Sheet.セル高さ調整
    Case "セル調整_両方"
      Call Ctl_Sheet.セル幅調整
      Call Ctl_Sheet.セル高さ調整
    Case "セル幅取得"
      Call Library.getColumnWidth
    
    'セル編集------------------------------------
    Case "Trim01"
        Call Ctl_Cells.Trim01
    Case "Trim02"
        Call Ctl_Cells.全空白削除
    Case "中黒点付与"
      Call Ctl_Cells.中黒点付与
    Case "連番追加"
      Call Ctl_Cells.連番追加
    Case "全半角変換"
      Call Ctl_Cells.英数字全半角変換
    Case "取り消し線"
      Call Ctl_Cells.取り消し線設定
    Case "コメント挿入"
      Call Ctl_Cells.コメント挿入
    Case "コメント削除"
      Call Ctl_Cells.コメント削除
    Case "コメント整形"
      Call Ctl_format.コメント整形
    
    Case "行例入れ替えて貼付け"
      Call Ctl_Cells.行例を入れ替えて貼付け
    Case "ゼロ埋め"
      Call Ctl_Cells.ゼロ埋め
    
    
    '数式編集------------------------------------
    Case "formula01"
      Call Ctl_Formula.formula01
    
    '整形------------------------------------
    Case "整形_1"
      Call Ctl_format.移動やサイズ変更をする
    Case "整形_2"
      Call Ctl_format.移動する
    Case "整形_3"
      Call Ctl_format.移動やサイズ変更をしない
    Case "余白ゼロ"
      Call Ctl_format.余白ゼロ
    
    '画像保存------------------------------------
    Case "saveImage"
      Call Ctl_Image.saveSelectArea2Image
    
    '罫線[クリア]--------------------------------
    Case "罫線_クリア"
      Call Library.罫線_クリア
    Case "罫線_クリア_中央線_横"
      Call Library.罫線_中央線削除_横
    Case "罫線_クリア_中央線_縦"
      Call Library.罫線_中央線削除_縦
    
    '罫線[表]------------------------------------
    Case "罫線_表_実線"
      Call Library.罫線_実線_格子
    Case "罫線_表_破線B"
      Call Library.罫線_表
    Case "罫線_表_破線C"
      Call Library.罫線_破線_格子
      Call Library.罫線_実線_水平
      Call Library.罫線_実線_囲み
    
    '罫線[破線]----------------------------------
    Case "罫線_破線_水平"
      Call Library.罫線_破線_水平
    Case "罫線_破線_垂直"
      Call Library.罫線_破線_垂直
    Case "罫線_破線_左"
      Call Library.罫線_破線_左
    Case "罫線_破線_右"
      Call Library.罫線_破線_右
    Case "罫線_破線_左右"
      Call Library.罫線_破線_左右
    Case "罫線_破線_上"
      Call Library.罫線_破線_上
    Case "罫線_破線_下"
      Call Library.罫線_破線_下
    Case "罫線_破線_上下"
      Call Library.罫線_破線_上下
    Case "罫線_破線_囲み"
      Call Library.罫線_破線_囲み
    Case "罫線_破線_格子"
      Call Library.罫線_破線_格子
    
    '罫線[実線]----------------------------------
    Case "罫線_実線_水平"
      Call Library.罫線_実線_水平
    Case "罫線_実線_垂直"
      Call Library.罫線_実線_垂直
    Case "罫線_実線_左右"
      Call Library.罫線_実線_左右
    Case "罫線_実線_上下"
      Call Library.罫線_実線_上下
    Case "罫線_実線_囲み"
      Call Library.罫線_実線_囲み
    Case "罫線_実線_格子"
      Call Library.罫線_実線_格子
    
    '罫線[二重線]----------------------------------
    Case "罫線_二重線_左"
      Call Library.罫線_二重線_左
    Case "罫線_二重線_左右"
      Call Library.罫線_二重線_左右
    Case "罫線_二重線_上"
      Call Library.罫線_二重線_上
    Case "罫線_二重線_下"
      Call Library.罫線_二重線_下
    Case "罫線_二重線_上下"
      Call Library.罫線_二重線_上下
    Case "罫線_二重線_囲み"
      Call Library.罫線_二重線_囲み
      
    'データ生成-----------------------------------
    Case "連番設定"
      Call Ctl_Cells.連番設定
    Case "連番生成"
      Call Ctl_Cells.連番追加
    Case "桁数固定数値"
      Call Ctl_sampleData.数値_桁数固定(Selection.count)
    Case "範囲指定数値"
      Call Ctl_sampleData.数値_範囲
    Case "姓"
      Call Ctl_sampleData.名前_姓(Selection.count)
    Case "名"
      Call Ctl_sampleData.名前_名(Selection.count)
    Case "氏名"
      Call Ctl_sampleData.名前_フルネーム(Selection.count)
    Case "日付"
      Call Ctl_sampleData.日付_日(Selection.count)
    Case "時間"
      Call Ctl_sampleData.日付_時間(Selection.count)
    Case "日時"
      Call Ctl_sampleData.日時(Selection.count)
    Case "文字"
      Call Ctl_sampleData.その他_文字(25)
    
    
    Case Else
      Call Library.showDebugForm("リボンメニューなし", control.ID, "Error")
      Call Library.showNotice("リボンメニューなし", control.ID, "Error")
  End Select
  
  '処理終了--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm("", , "end")
  Call init.unsetting
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

