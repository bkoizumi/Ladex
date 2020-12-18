Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook   As Workbook
Public targetBook As Workbook


'ワークシート用変数------------------------------
Public sheetsetting   As Worksheet
Public sheetNotice    As Worksheet
Public sheetStyle     As Worksheet
Public sheetTestData  As Worksheet
Public sheetRibbon    As Worksheet
Public sheetFavorite  As Worksheet


'グローバル変数----------------------------------
Public Const thisAppName = "BK_Library"
Public Const thisAppVersion = "0.0.4.0"

'レジストリ登録用サブキー
Public Const RegistryKey  As String = "BK_Library"
Public RegistrySubKey     As String
Public RegistryRibbonName As String

'設定値保持
Public setVal         As Object


'ファイル関連
Public logFile As String

'処理時間計測用
Public StartTime          As Date
Public StopTime           As Date



'リボン関連--------------------------------------
Public ribbonUI       As Office.IRibbonUI
Public ribbonVal      As Object


'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  
  On Error GoTo catchError
'  ThisWorkbook.Save

  'レジストリ関連設定------------------------------------------------------------------------------
  RegistrySubKey = "Main"
  RegistryRibbonName = "RP_" & ActiveWorkbook.Name
  
  If ThisBook Is Nothing Or reCheckFlg = True Then
  Else
    Exit Function
  End If

  'ブックの設定
  Set ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  Set sheetsetting = ThisBook.Worksheets("設定")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetStyle = ThisBook.Worksheets("Style")
  Set sheetTestData = ThisBook.Worksheets("testData")
  Set sheetRibbon = ThisBook.Worksheets("Ribbon")
  Set sheetFavorite = ThisBook.Worksheets("Favorite")

  
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
        
  '設定値読み込み----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  setVal.add "debugMode", "develop"
  
  Set ribbonVal = Nothing
  Set ribbonVal = CreateObject("Scripting.Dictionary")
  For line = 2 To sheetRibbon.Cells(Rows.count, 1).End(xlUp).Row
    If sheetRibbon.Range("A" & line) <> "" Then
      ribbonVal.add "Lbl_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("B" & line).Text
      ribbonVal.add "Act_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("C" & line).Text
      ribbonVal.add "Sup_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("D" & line).Text
      ribbonVal.add "Dec_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("E" & line).Text
      ribbonVal.add "Siz_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("F" & line).Text
      ribbonVal.add "Img_" & sheetRibbon.Range("A" & line).Text, sheetRibbon.Range("G" & line).Text
    End If
  Next
  
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  
End Function


