Attribute VB_Name = "Ctl_Ribbon"
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Function MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If


'**************************************************************************************************
' * リボンメニュー初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'読み込み時処理
Function onLoad(ribbon As IRibbonUI)
  Call init.setting
  
  Set ribbonUI = ribbon
  
'  Call Library.showDebugForm("ribbonUI" & "," & CStr(ObjPtr(ribbonUI)))
  Call Library.setRegistry("ribbonUI", CStr(ObjPtr(ribbonUI)))
  
  ribbonUI.ActivateTab ("BK_Library")
  ribbonUI.Invalidate
  
End Function


'==================================================================================================
'更新
Function Refresh()
  Call init.setting
  
  #If VBA7 And Win64 Then
    Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
  #Else
    Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
  #End If
  
  ribbonUI.ActivateTab ("BK_Library")
  ribbonUI.Invalidate
End Function
  
  
'==================================================================================================
'シート一覧メニュー
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  
  Call init.setting
   
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
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
      
      .SetAttribute "onAction", "BK_Library.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  
  'Call Library.showDebugForm(DOMDoc.XML)
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
End Function

'--------------------------------------------------------------------------------------------------
Function dMenuRefresh(control As IRibbonControl)
  
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
    #End If
  End If
  ribbonUI.Invalidate
End Function


'--------------------------------------------------------------------------------------------------
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


'--------------------------------------------------------------------------------------------------
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

'--------------------------------------------------------------------------------------------------
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


'--------------------------------------------------------------------------------------------------
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
   
'  On Error GoTo catchError
  Call init.setting
   
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menuの作成

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  If Workbooks.count = 0 Then
    endLine = 100
  Else
    endLine = sheetFavorite.Cells(Rows.count, 1).End(xlUp).row
  End If
  
  For line = 2 To endLine
    If sheetFavorite.Range("A" & line) <> "" Then
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "Favorite_" & line
        .SetAttribute "label", objFSO.GetFileName(sheetFavorite.Range("A" & line))
        .SetAttribute "imageMso", "Favorites"
        .SetAttribute "onAction", "BK_Library.xlam!Ctl_Ribbon.OpenFavoriteList"
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


'--------------------------------------------------------------------------------------------------
Function OpenFavoriteList(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  
  line = Replace(control.ID, "Favorite_", "")
  fileNamePath = sheetFavorite.Range("A" & line)
  
  If Library.chkFileExists(fileNamePath) Then
    Workbooks.Open fileName:=fileNamePath
  End If
  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function





'Label 設定----------------------------------------------------------------------------------------
Public Function getLabel(control As IRibbonControl, ByRef setRibbonVal)
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = Replace(ribbonVal("Lbl_" & control.ID), "<BR>", vbNewLine)
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
  setRibbonVal = ribbonVal("Act_" & control.ID)
  
  If setRibbonVal Like "*Ctl_Ribbon.*" Then
    Call Application.Run(setRibbonVal, control)
  
  ElseIf setRibbonVal = "" Then
    Call Library.showDebugForm("Act_" & control.ID)
  Else
    Call Application.Run(setRibbonVal)
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
  setRibbonVal = ribbonVal("Sup_" & control.ID)
End Function


'Description 設定----------------------------------------------------------------------------------
Public Function getDescription(control As IRibbonControl, ByRef setRibbonVal)
  Call init.setting
  setRibbonVal = Replace(ribbonVal("Dec_" & control.ID), "<BR>", vbNewLine)

End Function

'getImageMso 設定----------------------------------------------------------------------------------
Public Function getImage(control As IRibbonControl, ByRef image)
  On Error GoTo catchError
  
  Call init.setting
  image = ribbonVal("Img_" & control.ID)
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
  setRibbonVal = ribbonVal("Siz_" & control.ID)
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

'--------------------------------------------------------------------------------------------------
'有効/無効切り替え
Function getEnabled(control As IRibbonControl, ByRef returnedVal)
  Dim wb As Workbook
  Call init.setting
  
  If Workbooks.count = 0 Then
    returnedVal = False
  ElseIf setVal("debugMode") = "develop" Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'--------------------------------------------------------------------------------------------------
Function getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  returnedVal = Library.getRegistry("CustomRibbon")
End Function

'--------------------------------------------------------------------------------------------------
Function noDispTab(control As IRibbonControl)
  Call Library.setRegistry("CustomRibbon", False)
  Call RefreshRibbon
End Function

'--------------------------------------------------------------------------------------------------
Function setDispTab(control As IRibbonControl, pressed As Boolean)
  Call Library.setRegistry("CustomRibbon", pressed)
  Call RefreshRibbon
End Function


'--------------------------------------------------------------------------------------------------
Function RefreshRibbon()
  #If VBA7 And Win64 Then
    Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
  #Else
    Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
  #End If
  ribbonUI.Invalidate

End Function

'中央揃え------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function

