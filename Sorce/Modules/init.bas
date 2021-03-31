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
'Public RegistryRibbonName As String

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
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function usetting()

  Set ThisBook = Nothing
  
  'ワークシート名の設定
  Set sheetsetting = Nothing
  Set sheetNotice = Nothing
  Set sheetStyle = Nothing
  Set sheetTestData = Nothing
  Set sheetRibbon = Nothing
  Set sheetFavorite = Nothing

  '設定値読み込み
  Set setVal = Nothing
  Set ribbonVal = Nothing
End Function


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
  
  If ThisBook Is Nothing Or reCheckFlg = True Then
    Call usetting
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
  For line = 2 To sheetRibbon.Cells(Rows.count, 1).End(xlUp).row
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


'**************************************************************************************************
' * 名前定義
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError

  '名前の定義を削除
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" And _
      Not Name.Name Like "Slc*" And Not Name.Name Like "Pvt*" And Not Name.Name Like "Tbl*" Then
      Name.delete
    End If
  Next
  
  'VBA用の設定
  For line = 3 To sheetsetting.Cells(Rows.count, 1).End(xlUp).row
    If sheetsetting.Range("A" & line) <> "" Then
      sheetsetting.Range("B" & line).Name = sheetsetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  sheetsetting.Range("D3:D" & sheetsetting.Cells(Rows.count, 6).End(xlUp).row).Name = sheetsetting.Range("D2")
  

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

