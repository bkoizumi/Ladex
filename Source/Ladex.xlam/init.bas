Attribute VB_Name = "init"
Option Explicit


'ワークブック用変数------------------------------
Public BK_ThisBook          As Workbook
Public targetBook           As Workbook

'ワークシート用変数------------------------------
Public targetSheet          As Worksheet

Public targetRange          As Range


'グローバル変数----------------------------------
Public Const thisAppName    As String = "Ladex"
Public Const thisAppVersion As String = "1.4.4.0"
Public Const RelaxTools     As String = "Relaxtools.xlam"

Public funcName             As String
Public resetVal             As String
Public runFlg               As Boolean
Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
'Public LogLevel             As Long
Public arrFavCategory()
Public useStyle()
Public arrCells()



'レジストリ登録用キー----------------------------
Public Const RegistryKey    As String = "Ladex"
Public RegistrySubKey       As String


'設定値保持--------------------------------------
Public BK_setVal            As Object
Public sampleDataList       As Object
Public FrmVal               As Object


'ファイル/ディレクトリ関連-----------------------
Public logFile              As String
Public LadexDir             As String


'処理時間計測用----------------------------------
Public StartTime            As Date
Public StopTime             As Date



'リボン関連--------------------------------------
Public BK_ribbonUI          As Office.IRibbonUI
Public BK_ribbonVal         As Object
Public BKT_rbPressed        As Boolean

Public BKh_rbPressed        As Boolean
Public BKz_rbPressed        As Boolean
Public BKcf_rbPressed       As Boolean



'ユーザー関数関連--------------------------------
Public arryHollyday()       As Date

'ズーム関連--------------------------------------
Public defaultZoomInVal     As String





'**************************************************************************************************
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "init.unsetting"

  Set BK_ThisBook = Nothing
  
  '設定値読み込み
  Set BK_setVal = Nothing
  Set BK_ribbonVal = Nothing
  Set FrmVal = Nothing
  
  Set targetSheet = Nothing
  Set targetRange = Nothing
  
  Erase arrFavCategory
  Erase useStyle
  Erase arrCells
  
  logFile = ""
  LadexDir = ""
  
  If flg = True Then
    runFlg = False
  End If
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
  Dim tmpRegList
  
  Const funcName As String = "init.setting"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
'  ThisWorkbook.Save
'  If Workbooks.count = 0 Then
'    Call MsgBox("ブックが開かれていません", vbCritical, thisAppName)
'    Call Library.endScript
'    End
'  End If
  '----------------------------------------------

  If LadexDir = "" Or BK_setVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  'レジストリ関連
  RegistrySubKey = "Main"
  
  'ブックの設定
  Set BK_ThisBook = ThisWorkbook
  
  
  '設定値読み込み--------------------------------
  Set BK_setVal = Nothing
  Set BK_setVal = CreateObject("Scripting.Dictionary")
  
  endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 0 Then
    endLine = 11
  End If
  
  For line = 3 To endLine
    If LadexSh_Config.Range("A" & line) <> "" Then
      BK_setVal.add LadexSh_Config.Range("A" & line).Text, LadexSh_Config.Range("B" & line).Text
    End If
  Next
    
  'ユーザーフォームからの受け取り用--------------
  Set FrmVal = Nothing
  Set FrmVal = CreateObject("Scripting.Dictionary")
  FrmVal.add "commentVal", ""
  
  
  
  
  'レジストリ設定項目取得------------------------
  tmpRegList = GetAllSettings(thisAppName, "Main")
  For line = 0 To UBound(tmpRegList)
'    Debug.Print tmpRegList(line, 0) & "<-->" & tmpRegList(line, 1)
    BK_setVal.add tmpRegList(line, 0), tmpRegList(line, 1)
  Next
    
    
    
    
    
  'ログ出力設定----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  LadexDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex"
  logFile = LadexDir & "\log\ExcelMacro.log"
  Set wsh = Nothing
  
'  LogLevel = Split(BK_setVal("LogLevel"), ".")(0)
  
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function

'**************************************************************************************************
' * 名前定義
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  Const funcName As String = "init.名前定義"
  
  On Error GoTo catchError

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
  For line = 3 To LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
    If LadexSh_Config.Range("A" & line) <> "" Then
      LadexSh_Config.Range("B" & line).Name = LadexSh_Config.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  LadexSh_Config.Range("D3:D" & LadexSh_Config.Cells(Rows.count, 6).End(xlUp).Row).Name = LadexSh_Config.Range("D2")
  

  Exit Function
'エラー発生時------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function

