Attribute VB_Name = "Ctl_Sheet"
Option Explicit

'**************************************************************************************************
' * R1C1表記
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function R1C1表記()

  On Error Resume Next
  
  Call init.setting
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
  Else
    Application.ReferenceStyle = xlA1
  End If
  
End Function



'==================================================================================================
Function A1セル選択()
  Dim objSheet As Object
  Dim SheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  
  Const funcName As String = "Ctl_Sheet.A1セル選択"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    SheetName = objSheet.Name
    If Worksheets(SheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", SheetName, "debug")
      Application.GoTo Reference:=Worksheets(SheetName).Range("A1"), Scroll:=True
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, sheetCount + 1, sheetMaxCount + 1, SheetName & "A1セル選択")
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function すべて表示()
  Dim rowOutlineLevel As Long, colOutlineLevel As Long
  
  Const funcName As String = "Ctl_Sheet.すべて表示"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------

  If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
  End If
  If ActiveWindow.DisplayOutline = True Then
    ActiveSheet.Cells.ClearOutline
  End If
  ActiveSheet.Cells.EntireColumn.Hidden = False
  ActiveSheet.Cells.EntireRow.Hidden = False
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function 標準画面()
  Dim objSheet As Object
  Dim SheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim SelectAddress, setZoomLevel, resetBgColor, setGgridLine
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  PrgP_Max = 4
  PrgP_Cnt = 2
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  SelectAddress = Selection.Address
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  resetBgColor = Library.getRegistry("Main", "bgColor")
  setGgridLine = Library.getRegistry("Main", "gridLine")
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    SheetName = objSheet.Name
    If Worksheets(SheetName).Visible = True Then
      Worksheets(SheetName).Select
      
      '標準画面に設定
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.View = xlNormalView
      
      '表示倍率の指定
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      'ガイドラインの表示/非表示
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.DisplayGridlines = setGgridLine
  
      '背景白をなしにする
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      If resetBgColor = True Then
        With Application.FindFormat.Interior
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        With Application.ReplaceFormat.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
      End If
      'Application.Goto Reference:=Range("A1"), Scroll:=True
    End If
    Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  Range(SelectAddress).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
'==================================================================================================
Function シート管理_フォーム表示()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim addSheetName As String
  Const funcName As String = "Ctl_Sheet.シート管理_フォーム表示"
  
  Set targetBook = ActiveWorkbook
  Frm_Sheet.Show vbModeless
  
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

