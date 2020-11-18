Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook As Workbook
Public targetBook As Workbook


'ワークシート用変数------------------------------
Public sheetNotice As Worksheet
Public sheetStyle As Worksheet
Public sheetStyle2 As Worksheet
Public sheetRibbon As Worksheet


'グローバル変数----------------------------------
Public Const thisAppName = "BK_Library"
Public Const thisAppVersion = "0.0.4.0"

'レジストリ登録用サブキー
Public Const RegistryKey As String = "B.Koizumi"
Public Const RegistrySubKey As String = "BK_Library"

Public setVal As Collection

'ファイル関連
Public logFile As String


'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  
'  On Error GoTo catchError
  ThisWorkbook.Save

  If ThisBook Is Nothing Or reCheckFlg = True Then
  Else
    Exit Function
  End If

  'ブックの設定
  Set ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetStyle = ThisBook.Worksheets("Style")
  Set sheetStyle2 = ThisBook.Worksheets("Style2")
  Set sheetRibbon = ThisBook.Worksheets("Ribbon")

  Set setVal = New Collection
  With setVal
    .Add Item:="develop", Key:="debugMode"
  End With
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
  
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  
End Function


