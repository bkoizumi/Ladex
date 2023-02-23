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
Public Const thisAppName    As String = "Work Breakdown Structure for Excel"
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

Public Const startLine As Long = 7


'ファイル/ディレクトリ関連-----------------------
Public logFile              As String


'担当者情報--------------------------------------
Public lstAssign()          As String



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
  
  On Error GoTo catchError
  
  If logFile = "" Or setVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  Set targetBook = ActiveWorkbook
  
  'ワークシート名の設定
  Set Sh_PARAM = targetBook.Worksheets("PARAM")
  Set Sh_WBS = targetBook.Worksheets("WBS")
  
  'ログ出力設定----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  logFile = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex\log\WBS_ExcelMacro.log"
  Set wsh = Nothing
  
  
  
  '設定値読み込み--------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
'  setVal.Add item:="develop", Key:="debugMode"
  setVal.Add item:="5", Key:="LogLevel"
  
  endLine = Sh_PARAM.Cells(Rows.count, 1).End(xlUp).Row
  On Error Resume Next
  For line = 2 To endLine
    If Sh_PARAM.Range("A" & line) <> "" Then
      setVal.Add Sh_PARAM.Range("A" & line).Text, Sh_PARAM.Range("B" & line).Text
    End If
  Next
'  On Error GoTo catchError
  
  
'  Call WBS_Option.設定シートコピー("forAddin")
  Set sh_Option = ActiveWorkbook.Worksheets("Option")
  
  
  endLine = sh_Option.Cells(Rows.count, 1).End(xlUp).Row
  For line = 3 To endLine
    If sh_Option.Range("A" & line) <> "" Then
      setVal.Add sh_Option.Range("A" & line).Text, sh_Option.Range("B" & line).Text
    End If
  Next
  

  '担当者読み込み--------------------------------
  Set setAssign = Nothing
  Set setAssign = CreateObject("Scripting.Dictionary")
  
  endLine = sh_Option.Cells(Rows.count, 11).End(xlUp).Row
  On Error Resume Next
  For line = 4 To endLine
    If sh_Option.Range("K" & line) <> "" Then
      setAssign.Add sh_Option.Range("K" & line).Text, sh_Option.Range("K" & line).Interior.Color
    End If
  Next



  

  
  
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
  logFile = ""
End Function

'**************************************************************************************************
' * 休日設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHollyday(chkDate As Date, HollydayName As String)
  Dim line As Long, endLine As Long
  Dim strFilMessage() As Date
  
  '休日判定
  Call GetHollyday(CDate(chkDate), HollydayName)
  
  '土日を判定
  If HollydayName = "" Then
    If Weekday(chkDate) = vbSunday Then
      HollydayName = "Sunday"
    ElseIf Weekday(chkDate) = vbSaturday Then
      HollydayName = "Saturday"
    End If
  End If
  
  
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
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "担当者"

  endLine = setSheet.Cells(Rows.count, 17).End(xlUp).Row
  setSheet.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "休日リスト"

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'***********************************************************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function noDispSheet()

  Call init.setting
  'Worksheets("Help").Visible = xlSheetVeryHidden
  Worksheets("Tmp").Visible = xlSheetVeryHidden
  Worksheets("Notice").Visible = xlSheetVeryHidden
'  Worksheets("設定").Visible = xlSheetVeryHidden
  Worksheets("サンプル").Visible = xlSheetVeryHidden
  Worksheets(TeamsPlannerSheetName).Visible = xlSheetVeryHidden
  
  Worksheets(mainSheetName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("設定").Visible = True
  Worksheets("サンプル").Visible = True
  
  Worksheets(TeamsPlannerSheetName).Visible = True
  Worksheets(mainSheetName).Visible = True
  
  Worksheets(mainSheetName).Select
  
End Function





































