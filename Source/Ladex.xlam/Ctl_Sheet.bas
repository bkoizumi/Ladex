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

'**************************************************************************************************
' * セル幅・高さ調整
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function セル幅調整()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  
  Const funcName As String = "Ctl_Sheet.セル幅調整"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Cells.EntireColumn.AutoFit
  
  If IsNumeric(Range("A1").Text) Then
    Call Library.setColumnWidth
  Else
    For colLine = 1 To Columns.count
      If Cells(1, colLine).ColumnWidth > 30 Then
        colName = Library.getColumnName(colLine)
        Columns(colName & ":" & colName).ColumnWidth = 30
      End If
    Next
  End If
  Call Library.endScript(True)
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function セル高さ調整()
  Call Library.startScript
  Call init.setting
  
  Cells.EntireRow.AutoFit
  Call Library.endScript(True)
End Function

'==================================================================================================
Function A1セル選択()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
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
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      Application.Goto Reference:=Worksheets(sheetName).Range("A1"), Scroll:=True
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, sheetCount + 1, sheetMaxCount + 1, sheetName & "A1セル選択")
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
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
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting

  If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
  End If
  If ActiveWindow.DisplayOutline = True Then
    For rowOutlineLevel = 1 To 15
      DoEvents
      ActiveSheet.Outline.ShowLevels rowLevels:=rowOutlineLevel
      If Err.Number <> 0 Then
        Err.Clear
        Exit For
      End If
    Next
    
    For colOutlineLevel = 1 To 15
      DoEvents
      ActiveSheet.Outline.ShowLevels columnLevels:=colOutlineLevel
      If Err.Number <> 0 Then
        Err.Clear
        Exit For
      End If
    Next
  End If
  ActiveSheet.Cells.EntireColumn.Hidden = False
  ActiveSheet.Cells.EntireRow.Hidden = False
  
  Call Library.endScript(, True)
End Function

'==================================================================================================
Function 標準画面()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim SelectAddress, setZoomLevel, resetBgColor, setGgridLine
  
  'On Error Resume Next
  
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  SelectAddress = Selection.Address
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  resetBgColor = Library.getRegistry("Main", "bgColor")
  setGgridLine = Library.getRegistry("Main", "gridLine")
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    
    Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 0, 4, sheetName)
    
    If Worksheets(sheetName).Visible = True Then
      Worksheets(sheetName).Select
      
      '標準画面に設定
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 1, 4, sheetName)
      ActiveWindow.View = xlNormalView
      
      '表示倍率の指定
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 2, 4, sheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      'ガイドラインの表示/非表示
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 3, 4, sheetName)
      ActiveWindow.DisplayGridlines = setGgridLine
  
      '背景白をなしにする
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 4, 4, sheetName)
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
      Application.Goto Reference:=Range("A1"), Scroll:=True
    End If
    
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  Range(SelectAddress).Select
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  
End Function
