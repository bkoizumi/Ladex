Attribute VB_Name = "init"
Option Explicit

'ワークブック用変数------------------------------
Public LadexBook            As Workbook
Public targetBook           As Workbook

'ワークシート用変数------------------------------
Public targetSheet          As Worksheet

'セル用変数--------------------------------------
Public targetRange          As Range


'グローバル変数----------------------------------
Public Const thisAppName    As String = "Ladex"
Public Const thisAppVersion As String = "2.0.0.0"
Public Const RelaxTools     As String = "Relaxtools.xlam"
Public Const thisAppPasswd  As String = "Ladex"


Public funcName             As String
Public runFlg               As Boolean
Public G_LogLevel           As Long

Public arrFavCategory()
Public arrCells()

'プログレスバー関連------------------------------
Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
Public PbarCnt              As Long


'レジストリ登録用キー----------------------------
Public Const RegistryKey    As String = "Ladex"
Public RegistrySubKey       As String


'設定値保持--------------------------------------
Public dicVal               As Object
Public FrmVal               As Object
Public setIni               As Object
Public sampleDataList       As Object
Public resetVal             As String


'ファイル/ディレクトリ関連-----------------------
Public logFile              As String
Public LadexDir             As String
Public AddInDir             As String


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

'お気に入り関連----------------------------------
Public Const favoriteDebug  As Boolean = False

'セル関連----------------------------------------
Public Const maxColumnWidth As Long = 60
Public Const maxRowHeight   As Long = 200


'スタイル関連------------------------------------
Public useStyle()           As Variant
Public useStyleVal          As Object





'**************************************************************************************************
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "init.unsetting"

  '処理開始--------------------------------------
  On Error GoTo catchError
  '----------------------------------------------
  If flg = True Then
    Call resetGlobalVal
  End If
  
  Set LadexBook = Nothing
  
  '設定値読み込み
  Set dicVal = Nothing
  Set FrmVal = Nothing
  Set useStyleVal = Nothing
  
  Set targetSheet = Nothing
  Set targetRange = Nothing
  
  Erase arrFavCategory
  Erase useStyle
  Erase arrCells
  
  logFile = ""
  LadexDir = ""
  

  
  '処理終了--------------------------------------
  Exit Function
  '----------------------------------------------
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function resetGlobalVal()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "init.resetGlobalVal"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------

  '設定値読み込み
  Set dicVal = Nothing

  
  PrgP_Max = 2
  PrgP_Cnt = 0
  PbarCnt = 1
  runFlg = False
  
  '処理終了--------------------------------------
  Exit Function
  '----------------------------------------------
  
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
'  On Error GoTo catchError
  '----------------------------------------------

  If LadexDir = "" Or dicVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  'レジストリ関連
  RegistrySubKey = "Main"
  
  'ブックの設定
  Set LadexBook = ThisWorkbook
  
  'ログ出力設定----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  LadexDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex"
  logFile = LadexDir & "\log\ExcelMacro.log"
  AddInDir = wsh.SpecialFolders("AppData") & "\Microsoft\AddIns"

  
  Set wsh = Nothing
  Call Library.showDebugForm(funcName, , "function")
  
  If Library.Bookの状態確認 = True Then
    '設定値読み込み--------------------------------
    Set dicVal = Nothing
    Set dicVal = CreateObject("Scripting.Dictionary")
    
    endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
    If endLine = 0 Then
      endLine = 11
    End If
    
    For line = 3 To endLine
      If LadexSh_Config.Range("A" & line) <> "" Then
        dicVal.add LadexSh_Config.Range("A" & line).Text, LadexSh_Config.Range("B" & line).Text
      End If
    Next
    
    
    'ユーザーフォームからの受け取り用--------------
    Set FrmVal = Nothing
    Set FrmVal = CreateObject("Scripting.Dictionary")
    FrmVal.add "commentVal", ""
    
    'レジストリ設定項目取得------------------------
    tmpRegList = GetAllSettings(thisAppName, "Main")
    For line = 0 To UBound(tmpRegList)
      dicVal.add tmpRegList(line, 0), tmpRegList(line, 1)
    Next
    
    G_LogLevel = Split(dicVal("LogLevel"), ".")(0)
    
  Else
    Set FrmVal = Nothing
    Set FrmVal = CreateObject("Scripting.Dictionary")
    FrmVal.add "LogLevel", "5"
    G_LogLevel = 5
  
  End If
  

  
  
  '処理終了--------------------------------------
  Exit Function
  '----------------------------------------------
  
  
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

'==================================================================================================
Function resetsetVal()
  Dim line As Long, endLine As Long
  Dim tmpRegList
  
  Const funcName As String = "init.resetsetVal"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  '----------------------------------------------
  
  '設定値読み込み--------------------------------
  Set dicVal = Nothing
  Set dicVal = CreateObject("Scripting.Dictionary")
  
  endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 0 Then
    endLine = 11
  End If
  
  'レジストリ設定項目取得------------------------
  tmpRegList = GetAllSettings(thisAppName, "Main")
  For line = 0 To UBound(tmpRegList)
    dicVal.add tmpRegList(line, 0), tmpRegList(line, 1)
  Next
    
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function
