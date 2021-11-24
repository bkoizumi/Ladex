Attribute VB_Name = "init"
Option Explicit


'ワークブック用変数------------------------------
Public BK_ThisBook        As Workbook
Public targetBook         As Workbook


'ワークシート用変数------------------------------
Public targetSheet        As Worksheet

Public BK_sheetSetting    As Worksheet
Public BK_sheetNotice     As Worksheet
Public BK_sheetStyle      As Worksheet
Public BK_sheetTestData   As Worksheet
Public BK_sheetRibbon     As Worksheet
Public BK_sheetFavorite   As Worksheet
Public BK_sheetStamp      As Worksheet
Public BK_sheetHighLight  As Worksheet
Public BK_sheetHelp       As Worksheet
Public BK_sheetFunction   As Worksheet

'グローバル変数----------------------------------
Public Const thisAppName    As String = "Ladex"
Public Const thisAppVersion As String = "V1.0.0"
Public funcName             As String
Public resetVal             As String

Public Const RelaxTools     As String = "Relaxtools.xlam"



'レジストリ登録用サブキー
Public Const RegistryKey  As String = "Ladex"
Public RegistrySubKey     As String


'設定値保持
Public BK_setVal          As Object
Public sampleDataList     As Object


'ファイル/ディレクトリ関連
Public logFile            As String
Public LadexDir           As String


'処理時間計測用
Public StartTime          As Date
Public StopTime           As Date



'リボン関連--------------------------------------
Public BK_ribbonUI        As Office.IRibbonUI
Public BK_ribbonVal       As Object
Public BKT_rbPressed      As Boolean

Public BKh_rbPressed      As Boolean
Public BKz_rbPressed      As Boolean
Public BKcf_rbPressed     As Boolean



'ユーザー関数関連--------------------------------
Public arryHollyday()     As Date

'ズーム関連--------------------------------------
Public defaultZoomInVal   As String


'**************************************************************************************************
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetting()

  Set BK_ThisBook = Nothing
  
  'ワークシート名の設定
  Set BK_sheetSetting = Nothing
  Set BK_sheetNotice = Nothing
  Set BK_sheetStyle = Nothing
  Set BK_sheetTestData = Nothing
  Set BK_sheetRibbon = Nothing
  Set BK_sheetFavorite = Nothing

  '設定値読み込み
  Set BK_setVal = Nothing
  Set BK_ribbonVal = Nothing
  
  logFile = ""
End Function


'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
  
  On Error GoTo catchError
'  ThisWorkbook.Save
'  If Workbooks.count = 0 Then
'    Call MsgBox("ブックが開かれていません", vbCritical, thisAppName)
'    Call Library.endScript
'    End
'  End If

  'レジストリ関連設定------------------------------------------------------------------------------
  RegistrySubKey = "Main"
  
  If LadexDir = "" Or reCheckFlg = True Then
    Call init.unsetting
  Else
    Exit Function
  End If

  'ブックの設定
  Set BK_ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  Set BK_sheetSetting = BK_ThisBook.Worksheets("設定")
  Set BK_sheetNotice = BK_ThisBook.Worksheets("Notice")
  Set BK_sheetStyle = BK_ThisBook.Worksheets("Style")
  Set BK_sheetTestData = BK_ThisBook.Worksheets("testData")
'  Set BK_sheetRibbon = BK_ThisBook.Worksheets("Ribbon")
  Set BK_sheetFavorite = BK_ThisBook.Worksheets("Favorite")
  Set BK_sheetStamp = BK_ThisBook.Worksheets("Stamp")
  Set BK_sheetHighLight = BK_ThisBook.Worksheets("HighLight")
  Set BK_sheetHelp = BK_ThisBook.Worksheets("Help")
  Set BK_sheetFunction = BK_ThisBook.Worksheets("Function")
 
  
        
  '設定値読み込み----------------------------------------------------------------------------------
  Set BK_setVal = Nothing
  Set BK_setVal = CreateObject("Scripting.Dictionary")
  BK_setVal.add "debugMode", "develop"
  
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetSetting.Range("A" & line) <> "" Then
      BK_setVal.add BK_sheetSetting.Range("A" & line).Text, BK_sheetSetting.Range("B" & line).Text
    End If
  Next
  
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")

  LadexDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex"
  logFile = LadexDir & "\log\ExcelMacro.log"
  
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
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetSetting.Range("A" & line) <> "" Then
      BK_sheetSetting.Range("B" & line).Name = BK_sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  BK_sheetSetting.Range("D3:D" & BK_sheetSetting.Cells(Rows.count, 6).End(xlUp).Row).Name = BK_sheetSetting.Range("D2")
  

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

