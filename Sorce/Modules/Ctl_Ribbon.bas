Attribute VB_Name = "Ctl_Ribbon"
Private ctlEvent As New clsEvent

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
  On Error GoTo catchError
  
  
  Call init.setting
  
  BKh_rbPressed = Library.getRegistry("HighLight", ActiveWorkbook.Name)
  BKz_rbPressed = Library.getRegistry("ZoomIn", ActiveWorkbook.Name)
  
  Set BK_ribbonUI = ribbon
  
  Call Library.setRegistry("Main", "BK_ribbonUI", CStr(ObjPtr(BK_ribbonUI)))
  
  BK_ribbonUI.ActivateTab ("Ladex")
  BK_ribbonUI.Invalidate
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "リボンメニュー読込", True)
End Function


'==================================================================================================
' トグルボタンにチェックを設定する
Function HighLightPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKh_rbPressed
End Function

'==================================================================================================
' トグルボタンにチェックを設定する
Function ZoomInPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKz_rbPressed
End Function
'==================================================================================================
'更新
Function Refresh()
  On Error GoTo catchError
  
  Call init.setting
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  BK_ribbonUI.ActivateTab ("LiadexTab")
  BK_ribbonUI.Invalidate

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "リボンメニュー更新", True)
End Function

  
  
'==================================================================================================
'シート一覧メニュー
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  
  On Error GoTo catchError
  Call init.setting
  
  If BK_ribbonUI Is Nothing Then
    Stop
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

  For Each sheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = sheetName.Name
      .SetAttribute "id", "sheetID_" & sheetName.Index
      .SetAttribute "label", sheetName.Name
    
      If Sheets(sheetName.Name).Visible = True Then
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      ElseIf Sheets(sheetName.Name).Visible <> True Then
        .SetAttribute "imageMso", "SheetProtect"
      ElseIf ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
      End If
      
      .SetAttribute "onAction", "Liadex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  
  'Call Library.showDebugForm(DOMDoc.XML)
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
'  Call Library.showNotice(400, Err.Description, True)
End Function

'==================================================================================================
Function dMenuRefresh(control As IRibbonControl)
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  BK_ribbonUI.Invalidate
End Function


'==================================================================================================
Function selectActiveSheet(control As IRibbonControl)
  Dim sheetNameID As Integer
  Dim sheetCount As Integer
  Dim sheetName As Worksheet
  
  Call Library.startScript
  sheetNameID = Replace(control.ID, "sheetID_", "")
  
  If Sheets(sheetNameID).Visible <> True Then
    Sheets(sheetNameID).Visible = True
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
  
  Call Library.endScript
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


'==================================================================================================
#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If

  Dim ribbonObj As Object
  
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p)
End Function





' お気に入りメニュー作成---------------------------------------------------------------------------
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFSO As New FileSystemObject
   
  On Error GoTo catchError
  Call init.setting
  
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

  If Workbooks.count = 0 Then
    endLine = 100
  Else
    Call Ctl_Favorite.getList
    
    endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  End If
  
  For line = 2 To endLine
    If BK_sheetFavorite.Range("A" & line) <> "" Then
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "Favorite_" & line
        .SetAttribute "label", objFSO.GetFileName(BK_sheetFavorite.Range("A" & line))
        .SetAttribute "imageMso", "Favorites"
        .SetAttribute "onAction", "Liadex.xlam!Ctl_Ribbon.OpenFavoriteList"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
    End If
  Next
  DOMDoc.AppendChild Menu
  returnedVal = DOMDoc.XML
'  Call Library.showDebugForm(DOMDoc.XML)
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function OpenFavoriteList(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  
  line = Replace(control.ID, "Favorite_", "")
  fileNamePath = BK_sheetFavorite.Range("A" & line)
  
  If Library.chkFileExists(fileNamePath) Then
    Workbooks.Open fileName:=fileNamePath
  Else
    MsgBox "ファイルが存在しません" & vbNewLine & fileNamePath, vbExclamation
  End If
End Function





'Label 設定----------------------------------------------------------------------------------------
Public Function getLabel(control As IRibbonControl, ByRef setRibbonVal)
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = Replace(BK_ribbonVal("Lbl_" & control.ID), "<BR>", vbNewLine)
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function

'Action 設定---------------------------------------------------------------------------------------
Function getAction(control As IRibbonControl)
  Dim setRibbonVal As Variant
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = BK_ribbonVal("Act_" & control.ID)
  
  If setRibbonVal Like "*Ctl_Ribbon.*" Then
    Call Application.run(setRibbonVal, control)
  
  ElseIf setRibbonVal = "" Then
    Call Library.showDebugForm("Act_" & control.ID)
  Else
    Call Application.run(setRibbonVal)
  End If
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
  Call Library.endScript
End Function


'Supertip 設定-------------------------------------------------------------------------------------
Public Function getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  Call init.setting
  setRibbonVal = BK_ribbonVal("Sup_" & control.ID)
End Function


'Description 設定----------------------------------------------------------------------------------
Public Function getDescription(control As IRibbonControl, ByRef setRibbonVal)
  Call init.setting
  setRibbonVal = Replace(BK_ribbonVal("Dec_" & control.ID), "<BR>", vbNewLine)

End Function

'getImageMso 設定----------------------------------------------------------------------------------
Public Function getImage(control As IRibbonControl, ByRef image)
  On Error GoTo catchError
  
  Call init.setting
  image = BK_ribbonVal("Img_" & control.ID)
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function


'size 設定-----------------------------------------------------------------------------------------
Public Function getSize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = BK_ribbonVal("Siz_" & control.ID)
  Select Case setRibbonVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
      setRibbonVal = 0
  End Select
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function

'==================================================================================================
'有効/無効切り替え
Function getEnabled(control As IRibbonControl, ByRef returnedVal)
  Dim wb As Workbook
  Call init.setting
  
  If Workbooks.count = 0 Then
    returnedVal = False
  ElseIf BK_setVal("debugMode") = "develop" Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'==================================================================================================
Function getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
End Function

'==================================================================================================
Function noDispTab(control As IRibbonControl)
  Call Library.setRegistry("Main", "CustomRibbon", False)
  Call RefreshRibbon
End Function

'==================================================================================================
Function setDispTab(control As IRibbonControl, pressed As Boolean)
  Call Library.setRegistry("Main", "CustomRibbon", pressed)
  
  If pressed = True Then
    Call Refresh
  End If
  
End Function


'==================================================================================================
Function RefreshRibbon()
  On Error GoTo catchError

  #If VBA7 And Win64 Then
    Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
  #Else
    Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
  #End If
  BK_ribbonUI.Invalidate

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, Err.Description, True)
End Function





'**************************************************************************************************
' * リボンメニュー[オプション]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function settingImport(control As IRibbonControl)
  Call Main.設定_取込
End Function


'==================================================================================================
Function settingExport(control As IRibbonControl)
  Call Main.設定_抽出
End Function


'**************************************************************************************************
' * リボンメニュー[お気に入り]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
Function FavoriteAdd(control As IRibbonControl)

End Function

'==================================================================================================
Function FavoriteDetail(control As IRibbonControl)
  Call Ctl_Favorite.detail
End Function



'**************************************************************************************************
' * カスタマイズ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function defaultView(control As IRibbonControl)
  Call Main.標準画面
End Function

'==================================================================================================
Function HighLight(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKh_rbPressed = pressed
  
  Call init.setting
  Call Library.setRegistry("HighLight", ActiveWorkbook.Name, pressed)
  
  Call Ctl_HighLight.showStart(ActiveCell)
  If pressed = False Then
    'Call Library.unsetHighLight
'    Call Ctl_HighLight.showEnd
    Call Library.delRegistry("HighLight", ActiveWorkbook.Name)

  End If
End Function


'==================================================================================================
Function ZoomIn(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKz_rbPressed = pressed
  
  
  Call init.setting
  Call Library.setRegistry("ZoomIn", ActiveWorkbook.Name, pressed)
  
  If pressed = False Then
    Call Application.OnKey("{F2}")
    Call Library.delRegistry("ZoomIn", ActiveWorkbook.Name)
  Else
    Call Application.OnKey("{F2}", "Library.ZoomIn")
  End If
End Function


'==================================================================================================
Function delStyle(control As IRibbonControl)
  Call Main.スタイル削除
End Function




'**************************************************************************************************
' * リボンメニュー[罫線]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 罫線_削除(control As IRibbonControl)
  Call Library.罫線_クリア
End Function

'==================================================================================================
Function 罫線_表_破線(control As IRibbonControl)
  Call Library.罫線_表
End Function

'==================================================================================================
Function 罫線_表_実線(control As IRibbonControl)
  Call Library.罫線_実線_格子
End Function

'==================================================================================================
Function 罫線_破線_水平(control As IRibbonControl)
  Call Library.罫線_破線_水平
End Function

'==================================================================================================
Function 罫線_破線_垂直(control As IRibbonControl)
  Call Library.罫線_破線_垂直
End Function

'==================================================================================================
Function 罫線_破線_左右(control As IRibbonControl)
  Call Library.罫線_破線_左右
End Function

'==================================================================================================
Function 罫線_破線_上下(control As IRibbonControl)
  Call Library.罫線_破線_上下
End Function

'==================================================================================================
Function 罫線_破線_囲み(control As IRibbonControl)
  Call Library.罫線_破線_囲み
End Function

'==================================================================================================
Function 罫線_破線_格子(control As IRibbonControl)
  Call Library.罫線_破線_格子
End Function

'==================================================================================================
Function 罫線_二重線_左右(control As IRibbonControl)
  Call Library.罫線_二重線_左右
End Function

'==================================================================================================
Function 罫線_二重線_上下(control As IRibbonControl)
  Call Library.罫線_実線_上下
End Function

'==================================================================================================
Function 罫線_二重線_囲み(control As IRibbonControl)
  Call Library.罫線_二重線_囲み
End Function




'**************************************************************************************************
' * リボンメニュー[サンプルデータ生成]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function makeSampleData_SelectPattern(control As IRibbonControl)
  Call sampleData.パターン選択
End Function

'==================================================================================================
Function makeSampleData_DigitsInt(control As IRibbonControl)
  Call sampleData.数値_桁数固定
End Function

'==================================================================================================
Function makeSampleData_RangeInt(control As IRibbonControl)
  Call sampleData.数値_範囲
End Function

'==================================================================================================
Function makeSampleData_FamilyName(control As IRibbonControl)
  Call sampleData.名前_姓
End Function
'==================================================================================================
Function makeSampleData_Name(control As IRibbonControl)
  Call sampleData.名前_名
End Function

'==================================================================================================
Function makeSampleData_FullName(control As IRibbonControl)
  Call sampleData.名前_フルネーム
End Function





'**************************************************************************************************
' * リボンメニュー[その他]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setCenter(control As IRibbonControl)
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function
