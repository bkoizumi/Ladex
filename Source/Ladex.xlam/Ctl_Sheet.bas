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
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
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
  Dim sheetName As String, SetActiveSheet As String
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
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("SheetName", sheetName, "debug")
      
      Worksheets(sheetName).Select
      
      '標準画面に設定
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      ActiveWindow.View = xlNormalView
      
      '表示倍率の指定
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      'ガイドラインの表示/非表示
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      If setGgridLine = "表示しない" Then
        ActiveWindow.DisplayGridlines = False
      ElseIf setGgridLine = "表示する" Then
        ActiveWindow.DisplayGridlines = True
      ElseIf setGgridLine = "変更しない" Then
        'ActiveWindow.DisplayGridlines = setGgridLine
      End If
  
      '背景白をなしにする
      Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
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
      
      'A1を選択された状態にする
      Application.Goto Reference:=Range("A1"), Scroll:=True
      
      'RC表記からAQ表記へ変更
      If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
      End If
      
    End If
    Call Ctl_ProgressBar.showBar("標準画面設定", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
'  Range(SelectAddress).Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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
  
  '処理開始--------------------------------------
  Application.Cursor = xlWait
  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  Frm_Sheet.Show vbModeless
  
  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 幅設定()
  Dim SelectionCell As String

  
  Const funcName As String = "Ctl_Sheet.幅設定"

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
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  Cells.Select
  Range("A1").Activate
  Selection.ColumnWidth = Library.getRegistry("Main", "ColumnWidth")
  
  Range(SelectionCell).Select
  
  '処理終了--------------------------------------
  If runFlg = False Then
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
Function 高さ設定()
  Dim SelectionCell As String

  Const funcName As String = "Ctl_Sheet.高さ設定"

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
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  Cells.Select
  Range("A1").Activate
  Selection.rowHeight = Library.getRegistry("Main", "rowHeight")
  
  Range(SelectionCell).Select

  '処理終了--------------------------------------
  If runFlg = False Then
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
Function 体裁一括変更()
  Dim setGgridLine As String
  
  Const funcName As String = "Ctl_Sheet.体裁一括変更"

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
  '----------------------------------------------
  
  'オートシェイプ内の文字------------------------
  Dim slctObect As Shape
  On Error Resume Next
  For Each slctObect In ActiveSheet.Shapes
    slctObect.Select
    slctObect.Placement = xlMove
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = LadexsetVal("BaseFont")
      .TextRange.Font.NameFarEast = LadexsetVal("BaseFont")
      .TextRange.Font.Name = LadexsetVal("BaseFont")
      .TextRange.Font.Size = LadexsetVal("BaseFontSize")
      
      If .TextRange.Text <> "" Then
        .AutoSize = msoAutoSizeShapeToFitText
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        .WordWrap = msoTrue
      End If
    End With
    slctObect.Placement = xlFreeFloating
  Next
  On Error GoTo catchError
  
  'ガイドラインを非表示--------------------------
  setGgridLine = LadexsetVal("GridLine")
  If setGgridLine = "表示しない" Then
    ActiveWindow.DisplayGridlines = False
  ElseIf setGgridLine = "表示する" Then
    ActiveWindow.DisplayGridlines = True
  End If
      
  '高さ設定--------------------------------------
  Cells.Select
  Range("A1").Activate
  Selection.rowHeight = LadexsetVal("rowHeight")
  
  'フォント設定----------------------------------
  With Selection
    .Font.Name = LadexsetVal("BaseFont")
    .Font.Size = LadexsetVal("BaseFontSize")
    .VerticalAlignment = xlCenter
  End With
  
  '表示倍率--------------------------------------
  ActiveWindow.Zoom = LadexsetVal("ZoomLevel")
  
  Application.Goto Reference:=Range("A1"), Scroll:=True

  '処理終了--------------------------------------
  If runFlg = False Then
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
Function 指定フォントに設定()
  Dim SelectionCell As String

  Const funcName As String = "Ctl_Sheet.指定フォントに設定"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  
  'オートシェイプ内の文字------------------------
  Dim slctObect As Shape
  On Error Resume Next
  For Each slctObect In ActiveSheet.Shapes
    slctObect.Select
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = LadexsetVal("BaseFont")
      .TextRange.Font.NameFarEast = LadexsetVal("BaseFont")
      .TextRange.Font.Name = LadexsetVal("BaseFont")
      .TextRange.Font.Size = LadexsetVal("BaseFontSize")
      If .TextRange.Text <> "" Then
        .AutoSize = msoAutoSizeShapeToFitText
        .AutoSize = msoAutoSizeNone
      End If

    End With
  
  
  Next
  On Error GoTo catchError
  
  'フォント設定----------------------------------
  Cells.Select
  With Selection
    .Font.Name = LadexsetVal("BaseFont")
    .Font.Size = LadexsetVal("BaseFontSize")
    .VerticalAlignment = xlCenter
  End With

  Range(SelectionCell).Select
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function


'==================================================================================================
Function 連続シート追加()
  Dim sheetName As Variant
  
  Const funcName As String = "Ctl_Sheet.連続シート追加"

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
  
  Set FrmVal = Nothing
  Set FrmVal = CreateObject("Scripting.Dictionary")
  With Frm_Info
    .Caption = "連続シート生成"
    .TextBox.Value = ""
    .copySheet.Visible = True
    .Label1.Visible = True
    .Label2.Visible = True
    .Show
  End With
  
  Call Library.showDebugForm("copySheet", FrmVal("copySheet"), "debug")
  For Each sheetName In Split(FrmVal("SheetList"), vbNewLine)
    Call Library.showDebugForm("sheetName", sheetName, "debug")
    
    If Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") <> "≪新規シート≫" Then
      Worksheets(FrmVal("copySheet")).copy After:=Worksheets(Worksheets.count)
      ActiveSheet.Name = CStr(sheetName)
    
    ElseIf Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") = "≪新規シート≫" Then
      Worksheets.add(After:=Worksheets(Worksheets.count)).Name = CStr(sheetName)
    End If
    
    Application.Goto Reference:=Range("A1"), Scroll:=True
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
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
Function 画面枠固定()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim slctCell As String
  Const funcName As String = "Ctl_Sheet.画面枠固定"

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
  slctCell = ActiveCell.Address(False, False)
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True And Not (sheetName Like "《*》") Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      
      ActiveWindow.FreezePanes = False
      Range(slctCell).Select
      ActiveWindow.FreezePanes = True
      
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

