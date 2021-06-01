Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public BK_ThisBook   As Workbook
Public targetBook As Workbook


'ワークシート用変数------------------------------
Public BK_sheetsetting   As Worksheet
Public BK_sheetNotice    As Worksheet
Public BK_sheetStyle     As Worksheet
Public BK_sheetTestData  As Worksheet
Public BK_sheetRibbon    As Worksheet
Public BK_sheetFavorite  As Worksheet


'グローバル変数----------------------------------
Public Const thisAppName = "BK_Library"
Public Const thisAppVersion = "0.0.4.0"

'レジストリ登録用サブキー
'Public Const RegistryKey  As String = "BK_Library"
Public RegistrySubKey     As String
'Public RegistryRibbonName As String

'設定値保持
Public BK_setVal         As Object


'ファイル関連
Public logFile As String

'処理時間計測用
Public StartTime          As Date
Public StopTime           As Date



'リボン関連--------------------------------------
Public BK_ribbonUI       As Office.IRibbonUI
Public BK_ribbonVal      As Object


'**************************************************************************************************
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function usetting()

  Set BK_ThisBook = Nothing
  
  'ワークシート名の設定
  Set BK_sheetsetting = Nothing
  Set BK_sheetNotice = Nothing
  Set BK_sheetStyle = Nothing
  Set BK_sheetTestData = Nothing
  Set BK_sheetRibbon = Nothing
  Set BK_sheetFavorite = Nothing

  '設定値読み込み
  Set BK_setVal = Nothing
  Set BK_ribbonVal = Nothing
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
  
  If BK_ThisBook Is Nothing Or reCheckFlg = True Then
    Call usetting
  Else
    Exit Function
  End If

  'ブックの設定
  Set BK_ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  Set BK_sheetsetting = BK_ThisBook.Worksheets("設定")
  Set BK_sheetNotice = BK_ThisBook.Worksheets("Notice")
  Set BK_sheetStyle = BK_ThisBook.Worksheets("Style")
  Set BK_sheetTestData = BK_ThisBook.Worksheets("testData")
  Set BK_sheetRibbon = BK_ThisBook.Worksheets("Ribbon")
  Set BK_sheetFavorite = BK_ThisBook.Worksheets("Favorite")

  
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
        
  '設定値読み込み----------------------------------------------------------------------------------
  Set BK_setVal = Nothing
  Set BK_setVal = CreateObject("Scripting.Dictionary")
  BK_setVal.add "debugMode", "develop"
  
  Set BK_ribbonVal = Nothing
  Set BK_ribbonVal = CreateObject("Scripting.Dictionary")
  For line = 2 To BK_sheetRibbon.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetRibbon.Range("A" & line) <> "" Then
      BK_ribbonVal.add "Lbl_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("B" & line).Text
      BK_ribbonVal.add "Act_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("C" & line).Text
      BK_ribbonVal.add "Sup_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("D" & line).Text
      BK_ribbonVal.add "Dec_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("E" & line).Text
      BK_ribbonVal.add "Siz_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("F" & line).Text
      BK_ribbonVal.add "Img_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("G" & line).Text
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
  For line = 3 To BK_sheetsetting.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetsetting.Range("A" & line) <> "" Then
      BK_sheetsetting.Range("B" & line).Name = BK_sheetsetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  BK_sheetsetting.Range("D3:D" & BK_sheetsetting.Cells(Rows.count, 6).End(xlUp).Row).Name = BK_sheetsetting.Range("D2")
  

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

