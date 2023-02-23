Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook             As Workbook
Public targetBook           As Workbook

'ワークシート用変数------------------------------
Public targetSheet          As Worksheet
Public Sh_PARAM             As Worksheet
Public Sh_WBS               As Worksheet
Public sh_Sumally           As Worksheet
Public sh_Option            As Worksheet

'グローバル変数----------------------------------
Public Const thisAppName    As String = "test Case"
Public Const thisAppVersion As String = "1.0.0.0"

Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
Public runFlg               As Boolean
Public reCalflg             As Boolean
Public resetCellFlg         As Boolean

'設定値保持--------------------------------------
Public setVal               As Object
Public FrmVal               As Object
Public getVal               As Object
Public setAssign            As Object

Public resetVal             As String
Public SlctRange            As Range
Public PBarCnt              As Long


'Selenium関連------------------------------------
Public driver               As New Selenium.WebDriver
Public targetURL            As String
Public binPath              As String
Public BrowserProfilesDir   As String
Public waitFlg              As Boolean


Public resultArea1 As String
Public resultArea2 As String
Public resultArea3 As String
Public resultArea4 As String
Public resultArea5 As String



'ファイル/ディレクトリ関連-----------------------
Public logFile              As String

Public Const startLine As Long = 14

'***********************************************************************************************************************************************
' * 設定クリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Const funcName As String = "init.unsetting"
  
  If flg = True Then
    Call Library.showDebugForm("PrgP_Cnt", PrgP_Cnt, "debug")
    Call Library.showDebugForm("PrgP_Max", PrgP_Max, "debug")
  End If
  
  
  Set ThisBook = Nothing
  Set targetBook = Nothing
  
  Set targetSheet = Nothing
  Set Sh_PARAM = Nothing
  Set Sh_WBS = Nothing
  Set sh_Sumally = Nothing
  Set sh_Option = Nothing
  
  Set setVal = Nothing
  Set SlctRange = Nothing
  
  logFile = ""
  reCalflg = False
  PBarCnt = 1
  
  If flg = True Then
    PrgP_Cnt = 1
    PrgP_Max = 0
    
    runFlg = False
  End If
  
End Function
'***********************************************************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function setting(Optional reCheckFlg As Boolean = False)
  Dim line As Long
  Const funcName As String = "init.setting"
  
'  On Error GoTo catchError
  
  If Workbooks.count = 0 Then
    Exit Function
  End If
  
  If logFile = "" Or setVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If
  
  Set targetBook = ActiveWorkbook

  'ワークシート名の設定

  'ログ出力設定----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  logFile = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex\log\TestCase_ExcelMacro.log"
  binPath = wsh.SpecialFolders("AppData") & "\Bkoizumi\WebTools\bin\SeleniumBasic"
  BrowserProfilesDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\WebTools\BrowserProfiles"
  
  Set wsh = Nothing
  
  
  '設定値読み込み--------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")

  endLine = Sh_Config.Cells(Rows.count, 1).End(xlUp).Row
  On Error Resume Next
  For line = 2 To endLine
    If Sh_Config.Range("A" & line) <> "" Then
      setVal.Add Sh_Config.Range("A" & line).Text, Sh_Config.Range("B" & line).Text
    End If
  Next
  On Error GoTo catchError
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
  logFile = ""
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

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  'VBA用の設定
  For line = 3 To setSheet.Range("B4")
    If setSheet.Range("A" & line) <> "" Then
      setSheet.Range(setVal("cell_LevelInfo") & line).Name = setSheet.Range("A" & line)
    End If
  Next
  
  'ショートカットキーの設定
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).Row
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_ShortcutFuncName") & line) <> "" Then
      setSheet.Range(setVal("cell_ShortcutKey") & line).Name = setSheet.Range(setVal("cell_ShortcutFuncName") & line)
    End If
  Next
  
  
  endLine = setSheet.Cells(Rows.count, 11).End(xlUp).Row
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "Result"

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function
