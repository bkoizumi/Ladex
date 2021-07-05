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
  
  BKh_rbPressed = Library.getRegistry("Main", "HighLightFlg")
  BKz_rbPressed = Library.getRegistry("Main", "ZoomFlg")
  
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
' トグルボタンにチェックを設定する
Function confFormulaPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKcf_rbPressed
End Function

 
  
'==================================================================================================
'シート一覧メニュー
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  
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
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "sheetID_" & ActiveWorkbook.Name
      .SetAttribute "title", ActiveWorkbook.Name
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
  
  
  
  For Each sheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = sheetName.Name
      .SetAttribute "id", "sheetID_" & sheetName.Index
      .SetAttribute "label", sheetName.Name
    
      If ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
      ElseIf Sheets(sheetName.Name).Visible = True Then
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      ElseIf Sheets(sheetName.Name).Visible <> True Then
        .SetAttribute "imageMso", "SheetProtect"
      End If
      
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  
'  Call Library.showDebugForm(DOMDoc.XML)
  
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
  
  If BK_ribbonUI Is Nothing Then
  Else
    BK_ribbonUI.Invalidate
  End If
  
  
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
  
  If lRibbonPointer = 0 Then
    End
  End If
  
'  Stop
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p)
End Function





' お気に入りメニュー作成---------------------------------------------------------------------------
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
   
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

  Call Ctl_Favorite.getList
  If Workbooks.count = 0 Then
    endLine = 100
  Else
    endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  End If
  
  For line = 2 To endLine
    If BK_sheetFavorite.Range("A" & line) <> "" Then
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "Favorite_" & line
        .SetAttribute "label", objFso.GetFileName(BK_sheetFavorite.Range("A" & line))
        .SetAttribute "imageMso", "Favorites"
        .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.OpenFavoriteList"
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
  
  BKT_rbPressed = pressed
  Call Library.setRegistry("Main", "CustomRibbon", BKT_rbPressed)
  
  If pressed = True Then
    Call RefreshRibbon
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
'--------------------------------------------------------------------------------------------------
Function Optionshow(control As IRibbonControl)
  Ctl_Option.showOption
End Function

'--------------------------------------------------------------------------------------------------
Function OptionStyleImport(control As IRibbonControl)
  Call Ctl_Style.Import
End Function


'--------------------------------------------------------------------------------------------------
Function OptionStyleExport(control As IRibbonControl)
  Call Ctl_Style.Export
End Function

'--------------------------------------------------------------------------------------------------
Function OptionHighLight(control As IRibbonControl)
  Ctl_Option.HighLight
End Function

'--------------------------------------------------------------------------------------------------
Function OptionComment(control As IRibbonControl)
  Ctl_Option.Comment
End Function


'**************************************************************************************************
' * リボンメニュー[お気に入り]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'--------------------------------------------------------------------------------------------------
Function FavoriteAdd(control As IRibbonControl)
  Call Ctl_Favorite.addList
  
End Function

'--------------------------------------------------------------------------------------------------
Function FavoriteDetail(control As IRibbonControl)
  Call Ctl_Favorite.detail
End Function



'**************************************************************************************************
' * カスタマイズ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function defaultView(control As IRibbonControl)
  Call Main.標準画面
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect(control As IRibbonControl)
  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function

'--------------------------------------------------------------------------------------------------
Function defaultViewAndSave(control As IRibbonControl)
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  ActiveWorkbook.Save
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect2(control As IRibbonControl)
  Call Main.A1セル選択
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect2AndSave(control As IRibbonControl)
  Call Main.A1セル選択
  ActiveWorkbook.Save
End Function





'--------------------------------------------------------------------------------------------------
Function delStyle(control As IRibbonControl)
  Call Main.スタイル削除
End Function

'--------------------------------------------------------------------------------------------------
Function 名前定義削除(control As IRibbonControl)
  Call Main.名前定義削除
End Function

'--------------------------------------------------------------------------------------------------
Function 画像設定(control As IRibbonControl)
  'Call Main.すべて表示
End Function

'--------------------------------------------------------------------------------------------------
Function すべて表示(control As IRibbonControl)
  Call Main.すべて表示
End Function





'--------------------------------------------------------------------------------------------------
Function HighLight(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKh_rbPressed = pressed
  
  Call init.setting
  Call Library.setRegistry("Main", "HighLightFlg", pressed)
  
  Call Ctl_HighLight.showStart(ActiveCell)
  If pressed = False Then
    Call Library.delRegistry("Main", "HighLightFlg")

  End If
End Function

'--------------------------------------------------------------------------------------------------
Function dispR1C1(control As IRibbonControl)
  Call Main.R1C1表記
End Function


'--------------------------------------------------------------------------------------------------
Function AdjustWidth(control As IRibbonControl)
  Call Main.セル幅調整
End Function

'--------------------------------------------------------------------------------------------------
Function AdjustHeight(control As IRibbonControl)
  Call Main.セル高さ調整
End Function

'--------------------------------------------------------------------------------------------------
Function AdjustHeightAndWidth(control As IRibbonControl)
  Call Main.セル幅調整
  Call Main.セル高さ調整
  
End Function




'--------------------------------------------------------------------------------------------------
Function Zoom(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKz_rbPressed = pressed
  
  
  Call init.setting
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  If pressed = False Then
    Call Application.OnKey("{F2}")
    Call Library.delRegistry("Main", "ZoomFlg")
'    Call Ctl_DefaultVal.delVal("ZoomIn")
  Else
    Call Application.OnKey("{F2}", "Ctl_Zoom.ZoomIn")
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function Zoom01(control As IRibbonControl)
  Call Ctl_Zoom.Zoom01
  
End Function




'--------------------------------------------------------------------------------------------------
Function stamp01(control As IRibbonControl)
  Call Ctl_Stamp.押印_済印
End Function

'--------------------------------------------------------------------------------------------------
Function stamp02(control As IRibbonControl)
  Call Ctl_Stamp.押印_確認印
End Function

'--------------------------------------------------------------------------------------------------
Function stamp03(control As IRibbonControl)
  Call Ctl_Stamp.押印_済印
End Function



'--------------------------------------------------------------------------------------------------
Function confirmFormula(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKcf_rbPressed = pressed
  
  
  Call init.setting
  'Call Library.setRegistry("ZoomIn", ActiveWorkbook.Name, pressed)
  
  Call Ctl_Formula.数式確認
End Function

'--------------------------------------------------------------------------------------------------
Function formatComment(control As IRibbonControl)
  Call Main.コメント整形
End Function




'--------------------------------------------------------------------------------------------------
Function formula01(control As IRibbonControl)
  Call Ctl_Formula.formula01
End Function

'--------------------------------------------------------------------------------------------------
Function formula02(control As IRibbonControl)
  Call Main.xxxxxxxxxx
End Function

'--------------------------------------------------------------------------------------------------
Function formula03(control As IRibbonControl)
  Call Main.xxxxxxxxxx
End Function

'--------------------------------------------------------------------------------------------------
Function saveSelectArea2Image(control As IRibbonControl)
  Call Ctl_Image.saveSelectArea2Image
End Function

'**************************************************************************************************
' * リボンメニュー[文字編集]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function Trim01(control As IRibbonControl)
  Call Ctl_String.Trim01
  
End Function
'--------------------------------------------------------------------------------------------------
Function 中黒点付与(control As IRibbonControl)
  Call Ctl_String.中黒点付与
End Function

'--------------------------------------------------------------------------------------------------
Function 連番付与(control As IRibbonControl)
  Call Ctl_String.連番付与
End Function



'--------------------------------------------------------------------------------------------------
Function xxxxxxxxxx(control As IRibbonControl)
  Call Main.xxxxxxxxxx
End Function






'**************************************************************************************************
' * リボンメニュー[罫線]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function 罫線_クリア(control As IRibbonControl)
  Call Library.罫線_クリア
End Function
'--------------------------------------------------------------------------------------------------
Function 罫線_クリア_中央線_横(control As IRibbonControl)
  Call Library.罫線_中央線削除_横
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_クリア_中央線_縦(control As IRibbonControl)
  Call Library.罫線_中央線削除_縦
End Function


'**************************************************************************************************
' * リボンメニュー[罫線_表]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function 罫線_表_破線(control As IRibbonControl)
  Call Library.罫線_表
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_表_実線(control As IRibbonControl)
  Call Library.罫線_実線_格子
End Function


'**************************************************************************************************
' * リボンメニュー[罫線_破線]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function 罫線_破線_水平(control As IRibbonControl)
  Call Library.罫線_破線_水平
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_破線_垂直(control As IRibbonControl)
  Call Library.罫線_破線_垂直
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_破線_左右(control As IRibbonControl)
  Call Library.罫線_破線_左右
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_破線_上下(control As IRibbonControl)
  Call Library.罫線_破線_上下
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_破線_囲み(control As IRibbonControl)
  Call Library.罫線_破線_囲み
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_破線_格子(control As IRibbonControl)
  Call Library.罫線_破線_格子
End Function


'**************************************************************************************************
' * リボンメニュー[罫線_実線]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function 罫線_実線_水平(control As IRibbonControl)
  Call Library.罫線_実線_水平
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_実線_垂直(control As IRibbonControl)
  Call Library.罫線_実線_垂直
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_実線_左右(control As IRibbonControl)
  Call Library.罫線_実線_左右
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_実線_上下(control As IRibbonControl)
  Call Library.罫線_実線_上下
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_実線_囲み(control As IRibbonControl)
  Call Library.罫線_実線_囲み
End Function

'--------------------------------------------------------------------------------------------------
Function 罫線_実線_格子(control As IRibbonControl)
  Call Library.罫線_実線_格子
End Function




'**************************************************************************************************
' * リボンメニュー[罫線_二重線]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 罫線_二重線_左(control As IRibbonControl)
  Call Library.罫線_二重線_左
End Function
'==================================================================================================
Function 罫線_二重線_左右(control As IRibbonControl)
  Call Library.罫線_二重線_左右
End Function

'==================================================================================================
Function 罫線_二重線_下(control As IRibbonControl)
  Call Library.罫線_二重線_下
End Function

'==================================================================================================
Function 罫線_二重線_上下(control As IRibbonControl)
  Call Library.罫線_二重線_上下
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
'--------------------------------------------------------------------------------------------------
Function makeSampleData_SelectPattern(control As IRibbonControl)
  Call sampleData.パターン選択
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_DigitsInt(control As IRibbonControl)
  Call sampleData.数値_桁数固定
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_RangeInt(control As IRibbonControl)
  Call sampleData.数値_範囲
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_FamilyName(control As IRibbonControl)
  Call sampleData.名前_姓
End Function
'--------------------------------------------------------------------------------------------------
Function makeSampleData_Name(control As IRibbonControl)
  Call sampleData.名前_名
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_FullName(control As IRibbonControl)
  Call sampleData.名前_フルネーム
End Function





'**************************************************************************************************
' * リボンメニュー[その他]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function
