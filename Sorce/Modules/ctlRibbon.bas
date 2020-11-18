Attribute VB_Name = "ctlRibbon"
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If





'リボン関連--------------------------------------
Public ribbonUI As Office.IRibbonUI


'**************************************************************************************************
' * リボンメニュー初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'読み込み時処理------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  
  Call Library.setRegistry("RibbonPointer", CStr(ObjPtr(ribbonUI)))
  ribbonUI.ActivateTab ("BK_Library")
  
  'リボンの表示を更新する
  ribbonUI.Invalidate
End Function

'--------------------------------------------------------------------------------------------------
Function dMenu(control As IRibbonControl, ByRef returnedVal) ' 動的にメニューを作成する
   If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("RibbonPointer")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("RibbonPointer")))
    #End If
  End If
  ribbonUI.InvalidateControl "シート一覧"
  
  returnedVal = dynamicMenu()
End Function

'--------------------------------------------------------------------------------------------------
Function dMenuRefresh(control As IRibbonControl)
  
   If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("RibbonPointer")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("RibbonPointer")))
    #End If
  End If
  ribbonUI.InvalidateControl "シート一覧"
End Function

'--------------------------------------------------------------------------------------------------
Function dynamicMenu()
  Dim DOMDoc As Object ' Msxml2.DOMDocument
  Dim Menu As Object
  Dim Button As Object, subMenu As Object
  Dim sheetName As Worksheet
  Dim count As Long, maxCount As Long, menuCount As Long
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  
  Set Menu = DOMDoc.createElement("menu") ' menuの作成

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  For Each sheetName In ActiveWorkbook.Sheets
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", thisAppName & "_" & sheetName.Name
        .SetAttribute "label", sheetName.Name
        
      If Sheets(sheetName.Name).Visible = True Then
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      ElseIf Sheets(sheetName.Name).Visible <> True Then
        .SetAttribute "imageMso", "SheetProtect"
      
      End If
      If ActiveWorkbook.activeSheet.Name = sheetName.Name Then
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
      End If
        .SetAttribute "onAction", "activeSheet"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  dynamicMenu = DOMDoc.XML
'  Debug.Print DOMDoc.XML
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
End Function


'--------------------------------------------------------------------------------------------------
Function activeSheet(control As IRibbonControl)
  Dim sheetName As String
  
  Call Library.startScript
  sheetName = Replace(control.id, thisAppName & "_", "")
  
  If Sheets(sheetName).Visible <> True Then
    Sheets(sheetName).Visible = True
  End If
  Sheets(sheetName).Select
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  Call dynamicMenu
  Call Library.endScript
End Function


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
  p = 0: MoveMemory ribbonObj, p, LenB(p) '後始末
End Function


'Label 設定----------------------------------------------------------------------------------------
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 2)
End Sub


'Action 設定---------------------------------------------------------------------------------------
Sub getAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.id, 3)
  Application.run setRibbonVal
End Sub


'Supertip 設定-------------------------------------------------------------------------------------
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 5)
End Sub

'Description 設定----------------------------------------------------------------------------------
Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.id, 6)
End Sub

'getImageMso 設定----------------------------------------------------------------------------------
Public Sub getImageMso(control As IRibbonControl, ByRef image)
  setRibbonVal = getRibbonMenu(control.id, 7)
  image = setRibbonVal
End Sub


'size 設定-----------------------------------------------------------------------------------------
Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.id, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select
End Sub


'Ribbonシートから内容を取得------------------------------------------------------------------------
Function getRibbonMenu(menuId As String, offsetVal As Long)

  Dim getString As String
  Dim FoundCell As Range
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Call init.setting
  
  endLine = sheetRibbon.Cells(Rows.count, 1).End(xlUp).Row
  getRibbonMenu = Application.VLookup(menuId, sheetRibbon.Range("A2:G" & endLine), offsetVal, False)
  
  Call Library.endScript
  
  Exit Function
'エラー発生時=====================================================================================
catchError:
  getRibbonMenu = "エラー"
  Call Library.endScript
End Function



'--------------------------------------------------------------------------------------------------
Sub getVisible(control As IRibbonControl, ByRef returnedVal)
  returnedVal = Library.getRegistry("CustomRibbon")
End Sub

'--------------------------------------------------------------------------------------------------
Sub noDispTab(control As IRibbonControl)
  Call Library.setRegistry("CustomRibbon", False)
  Call RefreshRibbon
End Sub

'--------------------------------------------------------------------------------------------------
Sub setDispTab(control As IRibbonControl, pressed As Boolean)
  Call Library.setRegistry("CustomRibbon", pressed)
  Call RefreshRibbon
End Sub


'--------------------------------------------------------------------------------------------------
Function RefreshRibbon()
   If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("RibbonPointer")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("RibbonPointer")))
    #End If
  End If
  
  ribbonUI.Invalidate

End Function

'**************************************************************************************************
' * 独自リボンメニュー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'--------------------------------------------------------------------------------------------------
Function dispOption(control As IRibbonControl)
  Call Main.オプション画面表示
End Function


'--------------------------------------------------------------------------------------------------
Function setDefaultDisplay(control As IRibbonControl)
  Call Main.標準画面
End Function


'--------------------------------------------------------------------------------------------------
Function delStyle(control As IRibbonControl)
  Call Main.スタイル削除
End Function

'--------------------------------------------------------------------------------------------------
Function delVisibleNames(control As IRibbonControl)
  Call Main.名前定義削除
End Function


'--------------------------------------------------------------------------------------------------
Function setImage(control As IRibbonControl)
  Call Main.画像設定
End Function


'--------------------------------------------------------------------------------------------------
Function dispR1C1(control As IRibbonControl)
  Call Main.R1C1表記
End Function


'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function





'--------------------------------------------------------------------------------------------------
Function HighLight(control As IRibbonControl)
  Call Main.ハイライト
End Function


'--------------------------------------------------------------------------------------------------
Function lineSyle(control As IRibbonControl)
  
  
  Select Case control.id
    Case "xxxxxxxxxx"
    Case Else
  End Select
  
  
  
  
  
End Function
























