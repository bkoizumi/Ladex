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
  Call Library.showDebugForm(funcName, , "start")
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
  'Call Library.showDebugForm(funcName, , "start1")
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
  Dim targetBook  As Workbook
  Dim targetSheet As Worksheet
  
  Const funcName As String = "Ctl_Ribbon.HighLight"
  
  '処理開始--------------------------------------
  runFlg = True
  '  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
  BKh_rbPressed = pressed
  
  If pressed = False Then
      If Library.chkShapeName("HighLight_X", ActiveSheet) = True Then
        ActiveSheet.Shapes("HighLight_X").delete
      End If
      If Library.chkShapeName("HighLight_Y", ActiveSheet) = True Then
        ActiveSheet.Shapes("HighLight_Y").delete
      End If
    
    Call Library.setRegistry("Main", "HighLightFlg", pressed)
    Call Library.delRegistry("targetInfo", "HighLight_Book")
    Call Library.delRegistry("targetInfo", "HighLight_Sheet")
  
  Else
    Call Library.setRegistry("Main", "HighLightFlg", pressed)
    Call Library.setRegistry("targetInfo", "HighLight_Book", ActiveWorkbook.Name)
    Call Library.setRegistry("targetInfo", "HighLight_Sheet", ActiveSheet.Name)
    
    Call Ctl_HighLight.showStart(ActiveCell)
  End If
  Call Library.endScript
  
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
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
  BKz_rbPressed = pressed
  If pressed = False Then
'    Call Application.OnKey("{F2}")
    Call Library.delRegistry("Main", "ZoomFlg")
  Else
'    Call Application.OnKey("{F2}", "Ctl_Zoom.ZoomIn")
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
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  BKcf_rbPressed = pressed
  
  
  If BKcf_rbPressed = True Then
    Call Library.setRegistry("targetInfo", "Formula_Book", ActiveWorkbook.Name)
    Call Library.setRegistry("targetInfo", "Formula_Sheet", ActiveSheet.Name)
  Else
    Call Library.delRegistry("targetInfo", "Formula_Book")
    Call Library.delRegistry("targetInfo", "Formula_Sheet")
  End If
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
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  runFlg = True
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------
  
  fileNamePath = control.Tag
  
  If Library.chkFileExists(fileNamePath) Then
    If Library.chkBookOpened(fileNamePath) = True Then
      Call Library.showNotice(415, "", True)
    Else
      Select Case objFso.GetExtensionName(fileNamePath)
        Case "xls", "xlsx", "xlsm", "xlam"
          Workbooks.Open fileName:=fileNamePath
        Case Else
          CreateObject("Shell.Application").ShellExecute fileNamePath
      End Select
    End If
  Else
    Call Library.showNotice(404, fileNamePath, True)
  End If
  
  '処理終了--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'お気に入りファイル追加
Function FavoriteAddFile(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  Dim setCategory As Long

  Const funcName As String = "Ctl_Ribbon.FavoriteAddFile"

  '処理開始--------------------------------------
  On Error GoTo catchError
  runFlg = True
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------

  setCategory = Replace(control.ID, "M_FavoriteCategory", "")
  Call Library.showDebugForm("setCategory", setCategory, "debug")


  Call Ctl_Favorite.追加(setCategory, ActiveWorkbook.FullName)

  Call Library.delSheetData(LadexSh_Favorite)

  '処理終了--------------------------------------
  Call Library.showDebugForm(funcName, , "end")

  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'お気に入りファイル追加カテゴリーメニュー表示
Function FavoritesToAdd(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object, CategoryMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa, tmp, Category
  Dim categoryName As String, oldCategoryName As String
  
  Const funcName As String = "Ctl_Ribbon.FavoritesToAdd"

  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  If Library.Bookの状態確認 = False Then
    Call MsgBox("ブックが開かれていません", vbCritical, thisAppName)
    Call Library.errorHandle
    End
  End If
  
  Call Ctl_Favorite.リスト取得
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menuの作成

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
  With MenuSepa
    .setAttribute "id", "MS_お気に入り追加カテゴリー"
    .setAttribute "title", "お気に入り追加カテゴリー"
  End With
  Menu.appendChild MenuSepa
  Set MenuSepa = Nothing
  
  
  endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
  If IsEmpty(tmp) Then
    targetSheet.Range("A1") = "Category01"
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "M_FavoriteCategory" & 1
      .setAttribute "label", "Category01"
      .setAttribute "imageMso", "AddFolderToFavorites"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteAddFile"
    End With

    Menu.appendChild Button
    Set Button = Nothing
  Else
    Call Library.Sort_QuickSort(tmp, LBound(tmp), UBound(tmp), 0)
    oldCategoryName = ""
    line = 1
    
    For i = 0 To UBound(tmp)
      categoryName = Split(tmp(i, 0), "<L|>")(0)
      Call Library.showDebugForm("categoryName", categoryName, "debug")
      
      If oldCategoryName <> categoryName Then
        Set Button = DOMDoc.createElement("button")
        With Button
          .setAttribute "id", "M_FavoriteCategory" & line
          .setAttribute "label", categoryName
          .setAttribute "imageMso", "AddFolderToFavorites"
          .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteAddFile"
        End With
    
        Menu.appendChild Button
        Set Button = Nothing
        oldCategoryName = categoryName
        line = line + 1
      End If
    Next
    
  End If
  
  DOMDoc.appendChild Menu
  returnedVal = DOMDoc.XML
  'Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  Set CategoryMenu = Nothing
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  '処理終了--------------------------------------
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
' お気に入りメニュー
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object, CategoryMenu As Object, CategorymenuSeparator As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa, tmp, Category
  Dim FvrtCtgyCnt As Long, FvrtFileCnt As Long
  
  Const funcName As String = "Ctl_Ribbon.FavoriteMenu"

  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

'  Call Ctl_Favorite.getList
'  endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
  With MenuSepa
    .setAttribute "id", "FvrtCtgyList"
    .setAttribute "title", "カテゴリー一覧"
  End With
  Menu.appendChild MenuSepa
  Set MenuSepa = Nothing
  FvrtCtgyCnt = 1
  FvrtFileCnt = 1
  
  If Not IsEmpty(tmp) Then
    For line = 0 To UBound(tmp)
      Category = Split(tmp(line, 0), "<L|>")
      
      If Category(1) = 0 Then
        If line <> 0 Then
          Menu.appendChild CategoryMenu
        End If
        
        Set CategoryMenu = DOMDoc.createElement("menu")
        With CategoryMenu
          .setAttribute "id", "FvrtCtgy_" & FvrtCtgyCnt
          .setAttribute "label", Category(0)
          .setAttribute "imageMso", "AddFolderToFavorites"
        End With
        FvrtCtgyCnt = FvrtCtgyCnt + 1
        FvrtFileCnt = 1
      End If
    
      If tmp(line, 1) <> "" Then
        Set Button = DOMDoc.createElement("button")
        With Button
          '.setAttribute "id", Replace(tmp(line, 0), "<L|>", "_")
          .setAttribute "id", "FvrtCtgy_" & FvrtCtgyCnt - 1 & "_" & FvrtFileCnt
          .setAttribute "label", objFso.getFileName(tmp(line, 1))
          .setAttribute "tag", tmp(line, 1)
          
          'アイコンの設定
          Select Case objFso.GetExtensionName(tmp(line, 1))
            Case "xlsm", "xlsx"
              .setAttribute "imageMso", "MicrosoftExcel"
              
            Case "xlam"
              .setAttribute "imageMso", "FileSaveAsExcelXlsxMacro"
            
            Case "xls"
              .setAttribute "imageMso", "FileSaveAsExcel97_2003"
            
            Case "pdf"
              .setAttribute "imageMso", "FileSaveAsPdf"
            Case "docs"
              .setAttribute "imageMso", "FileSaveAsWordDocx"
            Case "text"
              .setAttribute "imageMso", "FileNewContext"
            Case "accdb"
              .setAttribute "imageMso", "MicrosoftAccess"
            Case "pptx"
              .setAttribute "imageMso", "MicrosoftPowerPoint"
            Case "csv"
              .setAttribute "imageMso", "FileNewContext"
            Case "html"
              .setAttribute "imageMso", "GroupWebPageNavigation"
            
            Case Else
              .setAttribute "imageMso", "FileNewContext"
              Call Library.showDebugForm("お気に入りアイコン", objFso.GetExtensionName(tmp(line, 1)), "warning")
          End Select
          
          
          '.setAttribute "supertip", tmp(line, 1)
          .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteFileOpen"
        End With
        CategoryMenu.appendChild Button
        Set Button = Nothing
        
        FvrtFileCnt = FvrtFileCnt + 1
      End If
    Next
    Menu.appendChild CategoryMenu

  Else
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "未登録"
      .setAttribute "label", "未登録"
      .setAttribute "imageMso", "FileNewContext"
      '.setAttribute "supertip", "未登録"
    End With
    Menu.appendChild Button
  End If
  DOMDoc.appendChild Menu
  returnedVal = DOMDoc.XML
  'Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
    
  Set CategoryMenu = Nothing
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
  Dim sheetName As Worksheet
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "シート管理"
      .setAttribute "title", "シート管理"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "シート管理表示"
      .setAttribute "label", "シート管理"
      .setAttribute "supertip", "シート管理"
      
      .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Menu.ladex_シート管理_フォーム表示"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
      With MenuSepa
        .setAttribute "id", "M_RelaxTools"
        .setAttribute "title", "RelaxToolsを利用"
      End With
      Menu.appendChild MenuSepa
      Set MenuSepa = Nothing

      Set Button = DOMDoc.createElement("button")
      With Button
        .setAttribute "id", "RelaxTools"
        .setAttribute "label", "RelaxTools"
        .setAttribute "supertip", "RelaxToolsのシート管理を起動"
        
        .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
        .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
      End With
      Menu.appendChild Button
      Set Button = Nothing
  End If
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "sheetID_0"
      .setAttribute "title", ActiveWorkbook.Name
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
  
  
  
  For Each sheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = sheetName.Name
      .setAttribute "id", "sheetID_" & sheetName.Index
      .setAttribute "label", sheetName.Name
    
      If ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
        .setAttribute "supertip", "アクティブシート"
        .setAttribute "imageMso", "ExcelSpreadsheetInsert"
        
      ElseIf Sheets(sheetName.Name).Visible = True Then
       '.SetAttribute "supertip", "アクティブシート"
        .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      
      ElseIf Sheets(sheetName.Name).Visible = 0 Then
        .setAttribute "supertip", "非表示シート"
        .setAttribute "imageMso", "SheetProtect"
      
      ElseIf Sheets(sheetName.Name).Visible = 2 Then
        .setAttribute "supertip", "マクロによる非表示シート"
        .setAttribute "imageMso", "ReviewProtectWorkbook"
      
      End If
      
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  Next

  DOMDoc.appendChild Menu
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
'RelaxTools
Function getRelaxTools(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa

  Const funcName As String = "Ctl_Ribbon.getRelaxTools"
  
  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"
  
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    'RelaxTools取得------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxToolsGet"
      .setAttribute "title", "RelaxToolの更新"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools_get"
      .setAttribute "label", "RelaxToolの更新"
      .setAttribute "image", "RelaxToolsLogo"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    'RelaxTools----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxTools"
      .setAttribute "title", "RelaxToolsを利用"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools01"
      .setAttribute "label", "シート管理"
      .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools02"
      .setAttribute "label", "書式リフレッシュ"
      '.SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools02"
    End With
    Menu.appendChild Button
    Set Button = Nothing

    'RelaxShapes---------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxShapes"
      .setAttribute "title", "RelaxShapesを利用"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes01"
      .setAttribute "label", "サイズ合わせ"
      .setAttribute "imageMso", "ShapesDuplicate"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes02"
      .setAttribute "label", "上位置合わせ"
      .setAttribute "imageMso", "ObjectsAlignTop"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes02"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes03"
      .setAttribute "label", "左位置合わせ"
      .setAttribute "imageMso", "ObjectsAlignLeft"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes03"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    'RelaxApps-----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxApps"
      .setAttribute "title", "RelaxAppsを利用"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxApps01"
      .setAttribute "label", "逆Ｌ罫線"
      .setAttribute "imageMso", "BorderDrawGrid"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxApps01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  Else
    'RelaxTools取得------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxTools"
      .setAttribute "title", "RelaxToolを入手"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools_get"
      .setAttribute "label", "RelaxToolを入手"
      .setAttribute "image", "RelaxToolsLogo"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  End If

  DOMDoc.appendChild Menu
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
  Dim sheetName As Worksheet
  Const funcName As String = "Ctl_Ribbon.selectActiveSheet"
  
  '処理開始--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "start1")
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
  For Each sheetName In ActiveWorkbook.Sheets
    If Sheets(sheetName.Name).Visible = True And sheetName.Name = Sheets(sheetNameID).Name Then
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
  Dim slctCells
  Dim slctCnt As Long
  Const funcName As String = "Ctl_Ribbon.setCenter"

  '処理開始--------------------------------------
  PrgP_Max = 2
  slctCnt = 1
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  For Each slctCells In Selection
    If TypeName(Selection) = "Range" Then
      Selection.HorizontalAlignment = xlCenterAcrossSelection
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, slctCnt, Selection.count, "")
      slctCnt = slctCnt + 1
    End If
  Next
  
  
  '処理終了--------------------------------------
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
    
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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
  PrgP_Max = 2
  PrgP_Cnt = 0

  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------
  
  Call Menu.各機能呼び出し(control.ID)

  
  '処理終了--------------------------------------
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end")
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

