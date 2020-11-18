Attribute VB_Name = "TestSpecification"
Public SetRowHeight As Long
Public SetAddRowHeight As Long
Public SetTestCaseFlg As Boolean
Public SetPrintOnePageRow As Long
Public SetFreezePanesCell As String
Public WindowZoomLevel As Long
Public SetTestCount As Long
Public SetActiveCell As String
Public SetActiveSheet As String
Public CheckRecal As Long
Public BeforeCloseFlg As Boolean

Public SetLineColorType As Boolean
Public SetLineColorColumn As String
Public SetLineColorValue As String
Public SetMonoChromePrinting  As Boolean




'***********************************************************************************************************************************************
' * 設定シート再設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_Reset()

  On Error Resume Next

  Worksheets("設定").Active
  
  ' 設定済の名前を削除
  Dim nm As Name
  For Each nm In ActiveWorkbook.Names
    nm.Delete
  Next nm

  ActiveWorkbook.Names.Add Name:="顧客名", RefersTo:=Range("E3")
  ActiveWorkbook.Names.Add Name:="作成日", RefersTo:=Range("E4")
  ActiveWorkbook.Names.Add Name:="作成者", RefersTo:=Range("E5")
  ActiveWorkbook.Names.Add Name:="プロジェクト名", RefersTo:=Range("E6")
  ActiveWorkbook.Names.Add Name:="システム名", RefersTo:=Range("E7")
  ActiveWorkbook.Names.Add Name:="表紙タイトル名称", RefersTo:=Range("E8")
  ActiveWorkbook.Names.Add Name:="ブラウザ", RefersTo:=Range("G3:G" & Cells(Rows.count, 7).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="結果", RefersTo:=Range("E13:E" & Cells(Rows.count, 5).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="動作確認環境", RefersTo:=Range("D13:D" & Cells(Rows.count, 4).End(xlUp).Row)
  
 
  SetLineColorFlg = False
  TestSpecification_SetShortcutKey          '独自ショートカットキー設定
'  TestSpecification_SetCellStyles           'セルのスタイル設定
  
'  TButton1 = Range("B3").Value
'  ribbonUI.Invalidate
End Sub


'***********************************************************************************************************************************************
' * ショートカットキー設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetShortcutKey()
  ' [F1]による『ヘルプ』画面の起動を無効にする。
  Call Application.OnKey("{F1}", "")
  

  If Worksheets("設定").Range("B16") <> "" Then
    Application.MacroOptions Macro:="FuncAddSheet", ShortcutKey:=Worksheets("設定").Range("B16")
  End If
  
  If Worksheets("設定").Range("B17") <> "" Then
    Application.MacroOptions Macro:="FuncSetNumber", ShortcutKey:=Worksheets("設定").Range("B17")
  End If

  If Worksheets("設定").Range("B18") <> "" Then
    Application.MacroOptions Macro:="FuncSetPrintArea", ShortcutKey:=Worksheets("設定").Range("B18")
  End If
  
  If Worksheets("設定").Range("B19") <> "" Then
    Application.MacroOptions Macro:="FuncSetResult", ShortcutKey:=Worksheets("設定").Range("B19")
  End If
  
  If Worksheets("設定").Range("B20") <> "" Then
    Application.MacroOptions Macro:="FuncSetResult", ShortcutKey:=Worksheets("設定").Range("B20")
  End If

End Function


'***********************************************************************************************************************************************
' * テスト結果報告書用環境設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_Init()

  Dim TmpTestCaseFlg As String
  
  On Error GoTo ErrHand
  
  ' 自動調整用の高さ
  SetRowHeight = Worksheets("設定").Range("B3")
  
  ' 自動調整用の高さ
  SetAddRowHeight = Worksheets("設定").Range("B4")
  
  ' 試験番号にA2のセル値を利用するかどうか
  TmpTestCaseFlg = Worksheets("設定").Range("B5")
  If TmpTestCaseFlg = "利用する" Then
    SetTestCaseFlg = True
  Else
    SetTestCaseFlg = False
  End If
  
  ' 印刷時、1ページに収める行数
  SetPrintOnePageRow = Worksheets("設定").Range("B6")
  SetFreezePanesCell = Worksheets("設定").Range("B7")
  
  ' 画面のズームレベル
  WindowZoomLevel = Worksheets("設定").Range("B8")
  
  '選択行の色付け
  If Range("B1") = "行・列とも" Then
    SetLineColorType = True
  Else
    SetLineColorType = False
  End If
  SetLineColorColumn = Range("B2")
  
  SetLineColorValue = Worksheets("設定").Range("B9").Interior.color
  
  'モノクロ印刷
  If Range("B10") = "する" Then
    SetMonoChromePrinting = True
  Else
    SetMonoChromePrinting = False
  End If
  
  
  
  ' マクロ実行前のアクティブシート、セル
  SetActiveSheet = activeSheet.Name
  
  If Selection.Address = "" Then
    SetActiveCell = "A1"
  Else
    SetActiveCell = Selection.Address
  End If
Exit Sub

ErrHand:
  
  SetActiveCell = "A1"
  Resume Next
  
End Sub

'***********************************************************************************************************************************************
' * テスト結果報告書用シート追加
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_AddSheet()

  Dim sheetName As String
  
  ' 入力用ボックスの表示
  sheetName = InputBox("機能名？", "機能名（シート名）入力", "")
  
  If sheetName <> "" Then
    Sheets("コピー用").Copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
    ActiveWorkbook.Sheets(Worksheets.count).Name = sheetName
    
    
  End If
  
End Function

'***********************************************************************************************************************************************
' * 試験番号設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetNumber()

  Dim RowCnt As Long
  Dim endLine As Long
  
  Dim TestCaseNo As Long
  
  Dim ColumLoopArray As Variant
  Dim ColumnLoopName As Variant
  Dim ColumnLoopCount As Long
 
  ' ======================= 処理開始 ======================
  Dim result As Boolean
   result = Library_CheckExcludeSheet(activeSheet.Name, 8)
  If result = False Then
    Exit Function
  End If
  

  Application.EnableCancelKey = xlErrorHandler
  On Error GoTo ErrHand
  
  
  ActiveWindow.DisplayGridlines = False
  TestSpecification_Init

  
  endLine = Cells(Rows.count, 5).End(xlUp).Row    ' アクティブシートの最終行取得
  
  TestCaseNo = 0
  
  ' プログレスバーの表示開始
  ProgressBar_ProgShowStart

  Columns("D:E").ColumnWidth = 50
  ' 試験番号生成 ====================================================
  For RowCnt = 9 To endLine
    TestCaseNo = TestCaseNo + 1
    
  ' 大項目のテストケース ============================================
    If Range("B" & RowCnt).Value <> "" Then
      TestSpecification_SetLineStyle1 (RowCnt)

  ' 中項目のテストケース ============================================
    ElseIf Range("C" & RowCnt).Value <> "" Then
      TestSpecification_SetLineStyle2 (RowCnt)

  ' 小項目のテストケース ============================================
    Else
      TestSpecification_SetLineStyle3 (RowCnt)
    End If

    ' 試験番号設定
    Range("A" & RowCnt).NumberFormatLocal = "@"
    If SetTestCaseFlg And Range("A2").Value <> "" Then
      Range("A" & RowCnt).Value = Range("A2").Value & "-" & Format(TestCaseNo, "000")
    Else
      Range("A" & RowCnt).Value = Format(TestCaseNo, "000")
    End If

    ' 背景色のクリア
    Range("A" & RowCnt & ":BD" & RowCnt).Select
      With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With

      With Selection.Font
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
      End With

    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
    ProgressBar_ProgShowCount "処理中", RowCnt, endLine, "1/5 試験番号設定：" & Range("A" & RowCnt).Value
    
    ' 高さの調節
    TestSpecification_setCellHeight (RowCnt)
  Next
  
  ' 最終行設定
  TestSpecification_SetEndLineStyle (RowCnt - 1)
  
  'テスト結果の集計
  ColumnLoopCount = 0

  
  ' 結果設定
  ColumLoopArray = Array("G", "Q", "AA", "AK", "AU")
  For Each ColumnLoopName In ColumLoopArray
    ProgressBar_ProgShowCount "処理中", ColumnLoopCount, UBound(ColumLoopArray), "2/5 結果設定中・・・・"
    Call TestSpecification_SetResultLineStyle(RowCnt - 1, ColumnLoopName)
    ColumnLoopCount = ColumnLoopCount + 1
  Next
  
  '確認日設定
  ColumLoopArray = Array("H", "K", "N", "R", "U", "X", "AB", "AE", "AH", "AL", "AO", "AR", "AV", "AY", "BB")
  ColumnLoopCount = 0
  For Each ColumnLoopName In ColumLoopArray
    ProgressBar_ProgShowCount "処理中", ColumnLoopCount, UBound(ColumLoopArray), "3/5 確認日設定中・・・・"
    Call TestSpecification_SetTestDay(RowCnt - 1, ColumnLoopName)
    ColumnLoopCount = ColumnLoopCount + 1
  Next
  
  '確認者設定
  ColumLoopArray = Array("I", "L", "O", "S", "V", "Y", "AC", "AF", "AI", "AN", "AP", "AS", "AW", "AZ", "BC")
  ColumnLoopCount = 0
  For Each ColumnLoopName In ColumLoopArray
    ProgressBar_ProgShowCount "処理中", ColumnLoopCount, UBound(ColumLoopArray), "4/5 確認者設定中・・・・"
    Call TestSpecification_SetTestUser(RowCnt - 1, ColumnLoopName)
    ColumnLoopCount = ColumnLoopCount + 1
  Next
  
  '備考設定
  ColumLoopArray = Array("J", "M", "P", "T", "W", "Z", "AD", "AF", "AJ", "AO", "AQ", "AT", "AX", "BA", "BD")
  ColumnLoopCount = 0
  For Each ColumnLoopName In ColumLoopArray
    ProgressBar_ProgShowCount "処理中", ColumnLoopCount, UBound(ColumLoopArray), "5/5 備考設定中・・・・"
    Call TestSpecification_SetTestComment(RowCnt - 1, ColumnLoopName)
    ColumnLoopCount = ColumnLoopCount + 1
  Next

'  TestSpecification_SetPrintArea


  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose
  
  ' ウィンドウ枠の固定
  If SetFreezePanesCell <> "" Then
    ActiveWindow.FreezePanes = False
    activeSheet.Outline.ShowLevels RowLevels:=2
    
    
    With ActiveWindow
      .ScrollRow = 1
      .ScrollColumn = 1
    End With
    
    Range("A1").Select
    Range(SetFreezePanesCell).Select
    ActiveWindow.FreezePanes = True
    activeSheet.Outline.ShowLevels RowLevels:=1
  End If
  
  '画面のズームレベル設定
  ActiveWindow.Zoom = WindowZoomLevel
  
  Erase ColumLoopArray
  
  
  ' ======================= 画面描写制御終了 ======================
  
  Range(SetActiveCell).Activate
    
Exit Function

' ======================= エラー発生時の処理 ======================
ErrHand:
  
'  Erase ColumLoopArray
  Range(SetActiveCell).Select
 
  ' プログレスバーの表示終了処理
  
  ProgressBar_ProgShowClose

  Call Library_ErrorHandle(Err.Number, Err.Description)
End Function

'***********************************************************************************************************************************************
' * 印刷範囲設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetPrintArea()

'  On Error GoTo ErrHand
  
  Dim endBookRowLine As Long
  Dim PageCnt As Long
  Dim W_PageNoCol As Long
  Dim RowCnt As Long
  
  TestSpecification_Init

' ======================= 画面描写制御開始 ======================
  endBookRowLine = Cells(Rows.count, 1).End(xlUp).Row
  W_PageNoCol = SetPrintOnePageRow + 4
  PageCnt = 1
  
  ' プログレスバーの表示開始
  ProgressBar_ProgShowStart
  
' ======================= 処理開始 ======================
  '改ページプレビュー
  ActiveWindow.View = xlPageBreakPreview
  
  'すべての改ページを解除
  activeSheet.ResetAllPageBreaks
  
  '印刷範囲をクリアする
  activeSheet.PageSetup.PrintArea = ""
  
  '印刷範囲の詳細設定
  With activeSheet.PageSetup
    .RightHeader = Range("A1").Value
    .CenterFooter = "&P / &N"
    .PrintTitleRows = "$5:$8"                 '行タイトル
    .PrintArea = "$A$5:$P$" & endBookRowLine
    .BlackAndWhite = SetMonoChromePrinting    '白黒印刷 True:する  False:しない
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False
  End With

  
  ActiveWindow.View = xlNormalView
  
' ======================= 画面描写制御終了 ======================
  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose
  Range(SetActiveCell).Select
  ActiveWindow.Zoom = WindowZoomLevel
  
Exit Sub

' ======================= エラー発生時の処理 ======================
ErrHand:

  ActiveWindow.View = xlNormalView
   
  ' プログレスバーの表示終了処理
  
  ProgressBar_ProgShowClose
  Call Library_ErrorHandle(Err.Number, Err.Description)

End Sub


'***********************************************************************************************************************************************
' * 大分類の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetLineStyle1(RowCnt As Long)

'  Range(Cells(RowCnt, 1), Cells(RowCnt, Cells(8, Columns.Count).End(xlToLeft).Column)).Select
  Range(Cells(RowCnt, 1), Cells(RowCnt, 56)).Select
  
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlDouble
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThick
  End With
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
'***********************************************************************************************************************************************
' * 中分類の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetLineStyle2(RowCnt As Long)

'  Range(Cells(RowCnt, 1), Cells(RowCnt, Cells(8, Columns.Count).End(xlToLeft).Column)).Select
  Range(Cells(RowCnt, 1), Cells(RowCnt, 56)).Select
  
  
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

  If Range("B" & RowCnt).Value = "" Then
    Range("B" & RowCnt).Borders(xlEdgeTop).LineStyle = xlNone
  End If
End Sub

'***********************************************************************************************************************************************
' * 小分類の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetLineStyle3(RowCnt As Long)

'  Range(Cells(RowCnt, 1), Cells(RowCnt, Cells(8, Columns.Count).End(xlToLeft).Column)).Select
  Range(Cells(RowCnt, 1), Cells(RowCnt, 56)).Select

  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlHairline
  End With
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  
  If Range("B" & RowCnt).Value = "" Then
    Range("B" & RowCnt).Borders(xlEdgeTop).LineStyle = xlNone
  End If
  If Range("C" & RowCnt).Value = "" Then
    Range("C" & RowCnt).Borders(xlEdgeTop).LineStyle = xlNone
  End If
  If Range("D" & RowCnt).Value = "" Then
    Range("D" & RowCnt).Borders(xlEdgeTop).LineStyle = xlNone
  End If
  
End Sub

'***********************************************************************************************************************************************
' * 最終行の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetEndLineStyle(RowCnt As Long)

  Range("A5:BD" & RowCnt).Select

  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With


  Range("A5:BD8").Select
  Range("BD8").Activate
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
End Sub


'***********************************************************************************************************************************************
' * 結果の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetResultLineStyle(RowCnt As Long, ColumnLoopName As Variant)

  Range(ColumnLoopName & "9:" & ColumnLoopName & RowCnt).Select
  With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
'    .MergeCells = False
  End With
  
  ' 入力規則の設定
  With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="=結果"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOff
    .ShowInput = True
    .ShowError = True
  End With
  
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlMedium
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End With
  
  Range("F" & RowCnt).Select
  With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    
    .Font.ColorIndex = xlAutomatic
    .Font.TintAndShade = 0
    .Font.Bold = False
  End With

End Function


'***********************************************************************************************************************************************
' * 確認日の罫線
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetTestDay(RowCnt As Long, ColumnLoopName As Variant)
  
  Range(ColumnLoopName & "9:" & ColumnLoopName & RowCnt).Select
  Selection.NumberFormatLocal = "yyyy/mm/dd"
  With Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
      .ColumnWidth = 12
  End With

  With Selection.Validation
      .Delete
      .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
      :=xlBetween
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = ""
      .ErrorTitle = ""
      .InputMessage = ""
      .ErrorMessage = ""
      .IMEMode = xlIMEModeOff
      .ShowInput = True
      .ShowError = True
  End With
End Function


'***********************************************************************************************************************************************
' * 確認者の入力規則
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetTestUser(RowCnt As Long, ColumnLoopName As Variant)
  
  Range(ColumnLoopName & "9:" & ColumnLoopName & RowCnt).Select
  With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  
  With Selection.Validation
      .Delete
      .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = ""
      .ErrorTitle = ""
      .InputMessage = ""
      .ErrorMessage = ""
      .IMEMode = xlIMEModeOn
      .ShowInput = True
      .ShowError = True
  End With
  
End Function


'***********************************************************************************************************************************************
' * 備考の入力規則
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetTestComment(RowCnt As Long, ColumnLoopName As Variant)

  Range(ColumnLoopName & "9:" & ColumnLoopName & RowCnt).Select
  With Selection.Validation
      .Delete
      .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
      :=xlBetween
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = ""
      .ErrorTitle = ""
      .InputMessage = ""
      .ErrorMessage = ""
      .IMEMode = xlIMEModeOn
      .ShowInput = True
      .ShowError = True
  End With
  
  Selection.NumberFormatLocal = "G/標準"
  With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  
  With Selection.Font
    .Name = "メイリオ"
    .Size = 8
  End With
End Function


'***********************************************************************************************************************************************
' * 行の高さの調節とテスト結果集計
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetResult()

  Dim RowCnt As Long
  Dim endLine As Long
  
  Dim TestCaseNo As Long
 
  ' ======================= 処理開始 ======================
  Dim result As Boolean
   result = Library_CheckExcludeSheet(activeSheet.Name, 9)
  If result = False Then
    Exit Function
  End If
  
'  On Error GoTo ErrHand
  
  TestSpecification_Init
  Range("A1") = activeSheet.Name
  
  endLine = Cells(Rows.count, 5).End(xlUp).Row    ' アクティブシートの最終行取得
  
  TestCaseNo = 0
  
  ' プログレスバーの表示開始
  ProgressBar_ProgShowStart
  Columns("D:E").ColumnWidth = 50
  
  ' 試験番号生成 ====================================================
  For RowCnt = 9 To endLine
    TestCaseNo = TestCaseNo + 1
    
  ' 大項目のテストケース ============================================
    If Range("B" & RowCnt).Value <> "" Then
      TestSpecification_SetLineStyle1 (RowCnt)

  ' 中項目のテストケース ============================================
    ElseIf Range("C" & RowCnt).Value <> "" Then
      TestSpecification_SetLineStyle2 (RowCnt)

  ' 小項目のテストケース ============================================
    Else
      TestSpecification_SetLineStyle3 (RowCnt)
    End If

    ' 試験番号設定
    Range("A" & RowCnt).NumberFormatLocal = "@"
    If SetTestCaseFlg And Range("A2").Value <> "" Then
      Range("A" & RowCnt).Value = Range("A2").Value & "-" & Format(TestCaseNo, "000")
    Else
      Range("A" & RowCnt).Value = Format(TestCaseNo, "000")
    End If

    ' 背景色のクリア
    Range("A" & RowCnt & ":BD" & RowCnt).Select
      With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With

      With Selection.Font
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
      End With

    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
    ProgressBar_ProgShowCount "処理中", RowCnt, endLine, "試験番号：" & Range("A" & RowCnt).Value
    
    
    ' 高さの調節
    Call TestSpecification_setCellHeight(RowCnt)

      
    'テスト結果の色付け ========================================================================================================
    Range("F" & RowCnt).Value = ""
    Range("F" & RowCnt).Value = TestSpecification_SetResultFlg(RowCnt)

  Next
  
  ' 最終行設定
  TestSpecification_SetEndLineStyle (RowCnt - 1)
  
  'テスト結果の集計
  Range("F1").NumberFormatLocal = "@"
  Range("F2").NumberFormatLocal = "@"
  Range("F3").NumberFormatLocal = "@"
  Range("F4").NumberFormatLocal = "@"
  
  Range("F1").Value = endLine - 8
  Range("F2").Value = WorksheetFunction.CountIf(Range("F9:F" & endLine), "NG") & "/" & _
                      WorksheetFunction.CountIf(Range("F9:F" & endLine), "OK.")
  
  Range("F3").Value = WorksheetFunction.CountIf(Range("F9:F" & endLine), "OK") + _
                      WorksheetFunction.CountIf(Range("F9:F" & endLine), "OK.")
                      
  Range("F4").Value = WorksheetFunction.CountIf(Range("F9:F" & endLine), "") & "/" & _
                                  WorksheetFunction.CountIf(Range("F9:F" & endLine), "対象外") + _
                                  WorksheetFunction.CountIf(Range("F9:F" & endLine), "準備作業")

  
  
  Dim ColumLoopArray As Variant
  Dim ColumnLoopName As Variant
  ColumLoopArray = Array("G", "Q", "AA", "AK", "AU")
  For Each ColumnLoopName In ColumLoopArray
    If Range(ColumnLoopName & "9").Value = "" And ColumnLoopName <> "G" Then
      Exit For
    End If
    
    Range(ColumnLoopName & "1").NumberFormatLocal = "@"
    Range(ColumnLoopName & "2").NumberFormatLocal = "@"
    Range(ColumnLoopName & "3").NumberFormatLocal = "@"
    Range(ColumnLoopName & "4").NumberFormatLocal = "@"
    
    Range(ColumnLoopName & "1").Value = endLine - 8
    Range(ColumnLoopName & "2").Value = WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG[致命的]") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG[重大]") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG[普通]") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG[限定的]") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "NG[軽微]") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "修正OK") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "確認OK")
    
    Range(ColumnLoopName & "3").Value = WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "OK") & "/ " & _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "修正OK") & "/ " & _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "確認OK")
                        
    Range(ColumnLoopName & "4").Value = WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "") & "/" & _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "対象外") + _
                                    WorksheetFunction.CountIf(Range(ColumnLoopName & "9:" & ColumnLoopName & endLine), "準備作業")

  Next

  If endLine <= 8 Then
    Range("F1").Value = 0
    Range("F2").Value = 0
    Range("F3").Value = 0
    Range("F4").Value = "0/0"
    For Each ColumnLoopName In ColumLoopArray
      Range(ColumnLoopName & "1").Value = 0
      Range(ColumnLoopName & "2").Value = 0
      Range(ColumnLoopName & "3").Value = "0/0/0"
      Range(ColumnLoopName & "4").Value = "0/0"
    Next
End If


  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose
  
  '画面のズームレベル設定
  Range("A1").Select
  ActiveWindow.Zoom = WindowZoomLevel
  Erase ColumLoopArray
  
  ' ======================= 画面描写制御終了 ======================
  
  Range(SetActiveCell).Activate
    
Exit Function

' ======================= エラー発生時の処理 ======================
ErrHand:
  
  Erase ColumLoopArray
  
  Range(SetActiveCell).Select
  
  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose

  Call Library_ErrorHandle(Err.Number, Err.Description)
End Function


'***********************************************************************************************************************************************
' * 総合結果判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetResultFlg(RowCnt As Long)

  Dim CheckFlg As String
  Dim ResultValue As String

  CheckFlg = False

  If Range("AU" & RowCnt).Value <> "" Then
    ResultValue = Range("AU" & RowCnt).Value
  ElseIf Range("AK" & RowCnt).Value <> "" Then
    ResultValue = Range("AK" & RowCnt).Value
  ElseIf Range("AA" & RowCnt).Value <> "" Then
    ResultValue = Range("AA" & RowCnt).Value
  ElseIf Range("Q" & RowCnt).Value <> "" Then
    ResultValue = Range("Q" & RowCnt).Value
  ElseIf Range("G" & RowCnt).Value <> "" Then
    ResultValue = Range("G" & RowCnt).Value
  End If
  
  Range("F" & RowCnt & ":F" & RowCnt).Select
  
  Select Case ResultValue
    Case "対象外", "準備作業"
      Range("A" & RowCnt & ",D" & RowCnt & ":BD" & RowCnt).Select
      
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E13").Interior.color  'RGB(255, 204, 204)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E13").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "対象外"
        
    Case "NG[致命的]"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E15").Interior.color   'RGB(255, 0, 0)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E15").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "NG"
        
    Case "NG[重大]"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E16").Interior.color  'RGB(255, 51, 51)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E16").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "NG"
        
    Case "NG"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E17").Interior.color  'RGB(255, 102, 102)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E17").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "NG"
        
    Case "NG[限定的]"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E18").Interior.color  'RGB(255, 153, 153)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E18").Font.color
        .TintAndShade = 0
      End With
     CheckFlg = "NG"

    Case "NG[軽微]"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E19").Interior.color  'RGB(255, 204, 204)
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E19").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "NG"
    

    Case "OK"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E14").Interior.color
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E14").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "OK"
    
    Case "修正OK"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E20").Interior.color
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E20").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "OK."
    
    Case "確認OK"
      With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = Worksheets("設定").Range("E21").Interior.color
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      With Selection.Font
        .color = Worksheets("設定").Range("E21").Font.color
        .TintAndShade = 0
      End With
      CheckFlg = "OK."
    
    Case ""
      With Range("F" & RowCnt).Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
      End With
      With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      CheckFlg = ""
      
      
    Case Else
      With Range("F" & RowCnt).Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
      End With
      With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      CheckFlg = "OK"
  End Select
  TestSpecification_SetResultFlg = CheckFlg
End Function


'***********************************************************************************************************************************************
' * バグ曲線生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_MakeGompertzCurve()

  Dim endLine As Long
  Dim sheetName As Object
  Dim result As Boolean
  Dim GrafCount As Long

  Dim SheetCount As Long
  Dim SheetCountOK As Long
  Dim SheetCountNG As Long
  Dim SheetCountNotTest As Long
  Dim SheetCountNotTestStr As String

  Dim AllCount As Long
  Dim AllCountOK As Long
  Dim AllCountNG As Long
  Dim AllCountNotTest As Long
  
  Dim TesterLineCount As Long
  
  Dim ColumLoopCount As Long
  Dim ColumnLoopName As String
      
  ' ======================= 処理開始 ======================
'  On Error GoTo ErrHand
  GrafCount = 0
  
  TestSpecification_Init
  
  Worksheets("総合結果").Select
  Worksheets("総合結果").Range("A1").Select
  TesterLineCount = 33
  
  
  If Not BeforeCloseFlg Then
    CheckRecal = MsgBox("再計算を行いますか？", vbYesNo + vbQuestion, "確認")
  Else
    CheckRecal = vbYes
  End If
  
  For Each sheetName In ActiveWorkbook.Sheets
    
    ' 結果計算の再実行
    result = Library_CheckExcludeSheet(sheetName.Name, 9)
    
    If result = True Then
      Worksheets(sheetName.Name).Activate
      Range("A1") = activeSheet.Name
      
      ' ======================= テスター情報抽出 ======================
      Worksheets("総合結果").Activate
      With Worksheets("総合結果").Range("H" & TesterLineCount)
        .Value = sheetName.Name
        .Select
        .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName.Name & "!" & "A9"
        .Font.color = RGB(0, 0, 0)
        .Font.Underline = False
        .Font.Size = 10
        .Font.Name = "Meiryo UI"
      End With


      Worksheets(sheetName.Name).Activate
      For ColumLoopCount = 13 To Worksheets("総合結果").Cells(32, Columns.count).End(xlToLeft).Column Step 1
        ColumnLoopName = Library_getColumnName(ColumLoopCount)
        If Worksheets("総合結果").Range(ColumnLoopName & "32").Value <> "" Then
          Worksheets("総合結果").Range(ColumnLoopName & TesterLineCount).Value = _
              WorksheetFunction.CountIf(Range("I:I"), Worksheets("総合結果").Range(ColumnLoopName & "32").Value)
        End If
      Next
      
      ' ======================= バグ曲線情報抽出 ======================
      If CheckRecal = vbYes Then
        TestSpecification_SetResult
      End If
      SheetCount = Range("F1").Value
      tmp = Split(Range("F2").Value, "/")
      SheetCountNG = val(tmp(0)) + val(tmp(1))
      SheetCountNGOK = val(tmp(1))

      SheetCountOK = Range("F3").Value
      SheetCountNotTestStr = Range("F4").Value
      tmp = Split(SheetCountNotTestStr, "/")
      SheetCountNotTest = val(tmp(0))
      SheetCountNotTestCase = val(tmp(1))
    
      AllCount = AllCount + SheetCount
      AllCountOK = AllCountOK + SheetCountOK
      AllCountNG = AllCountNG + SheetCountNG
      AllCountNotTest = AllCountNotTest + SheetCountNotTest
      AllCountNotTestCase = AllCountNotTestCase + SheetCountNotTestCase
        
    
      'テストケース総数
      Worksheets("総合結果").Range("I" & TesterLineCount).Value = SheetCountOK
      Worksheets("総合結果").Range("J" & TesterLineCount).Value = SheetCountNG
      Worksheets("総合結果").Range("K" & TesterLineCount).Value = SheetCountNotTest
      Worksheets("総合結果").Range("L" & TesterLineCount).Value = SheetCount
      
'      If SheetCount = (SheetCountOK + SheetCountNotTest) Then
'        Worksheets("総合結果").Activate
'        Range("H" & TesterLineCount & ":" & ColumnLoopName & TesterLineCount).Select
'        With Selection.Interior
'          .Color = RGB(222, 222, 222)
'        End With
'      End If
      
      '日別NG数計算
'      Public SheetNGCountRuikei As Object
'      Dim TestCount As Long
'      Dim ColumLoopArray As Variant
'      Dim ColumnLoopName As Variant
'
'      ColumLoopArray = Array("H", "R", "", "", "")
'      Set SheetNGCountRuikei = CreateObject("Scripting.Dictionary")
'      Endline = Cells(Rows.Count, 1).End(xlUp).Row
'      TestCount = Worksheets("総合結果").Range("B31").Value
  
      TesterLineCount = TesterLineCount + 1
    End If
  Next
  
  Worksheets("総合結果").Select
    
  '選択セルの行背景設定
  Call Library_SetLineColor("H33:" & ColumnLoopName & TesterLineCount, True, SetLineColorValue)
  
  '合計算出
  Range("H" & TesterLineCount) = "合計"
  For ColumLoopCount = 9 To Worksheets("総合結果").Cells(32, Columns.count).End(xlToLeft).Column Step 1
    ColumnLoopName = Library_getColumnName(ColumLoopCount)
    If ColumnLoopName <> "L" Then
      Range(ColumnLoopName & TesterLineCount) = "=SUM(" & ColumnLoopName & "33:" & ColumnLoopName & TesterLineCount - 1 & ")"
    End If
  Next
  
  '罫線
  Range("H32" & ":" & ColumnLoopName & TesterLineCount).Select
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlHairline
  End With
  
  ' 合計行の罫線
  Range("H" & TesterLineCount & ":" & ColumnLoopName & TesterLineCount).Select

  With Selection.Borders(xlEdgeTop)
    .LineStyle = xlDouble
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThick
  End With

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  If Range("A" & endLine).Value = "実施日" Then
    endLine = endLine + 1
  End If
  
  'OKが1つもなければ最終行を上書き
  If (AllCountOK < 0 And AllCountNG < 0) Then
    endLine = endLine
  ElseIf Format(Range("A" & endLine).Value, "yyyy/mm/dd") <> Format(Date, "yyyy/mm/dd") Then
    endLine = endLine + 1
  
  End If
  
  Range("A" & endLine).Value = Format(Date, "yyyy/mm/dd")
  Range("A" & endLine).NumberFormatLocal = "mm/dd"
  Range("B" & endLine).Value = AllCount
  Range("C" & endLine).Value = AllCountOK
  Range("D" & endLine).Value = AllCountNG
  Range("E" & endLine).Value = AllCountNotTest
  Range("F" & endLine).Value = AllCountNotTestCase
  

  Range("A" & endLine & ":" & "F" & endLine).Select
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ColorIndex = xlAutomatic
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = xlAutomatic
      .TintAndShade = 0
      .Weight = xlHairline
  End With
  With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = xlAutomatic
      .TintAndShade = 0
      .Weight = xlHairline
  End With
  With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ColorIndex = xlAutomatic
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ColorIndex = xlAutomatic
      .TintAndShade = 0
      .Weight = xlThin
  End With
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone



  '-----------------------------------------------------------------------------------------
  ' アクティブシート上に既存のグラフがあれば削除
  If activeSheet.ChartObjects.count > 0 Then
    For i = 1 To activeSheet.ChartObjects.count
      ' グラフ名が一致するか
      If activeSheet.ChartObjects(i).Name = "GompertzCurve" Then
        activeSheet.ChartObjects(i).Delete
        Exit For
      End If
    Next i
  End If
  
  Set chartObj = activeSheet.ChartObjects.Add(20, 20, 1000, 500)
  chartObj.Name = "GompertzCurve"
    With chartObj.Chart
      .ChartType = xlXYScatterSmoothNoMarkers
'      .Axes(xlValue).MajorUnit = 10

      .Axes(xlValue).HasMajorGridlines = True
      .Axes(xlValue).MajorGridlines.Border.LineStyle = xlContinuous
      .Axes(xlValue).MajorGridlines.Border.color = RGB(144, 136, 136)
      
      .Axes(xlValue).HasMinorGridlines = True
      .Axes(xlValue).MinorGridlines.Border.LineStyle = xlDot
      .Axes(xlValue).MinorGridlines.Border.Weight = xlHairline
      .Axes(xlValue).MinorGridlines.Border.color = RGB(222, 222, 222)
      
      .Axes(xlCategory).HasMajorGridlines = True
      .Axes(xlCategory).MajorGridlines.Border.LineStyle = xlContinuous
      .Axes(xlCategory).MinorGridlines.Border.color = RGB(144, 136, 136)

      .Axes(xlCategory).HasMinorGridlines = True
      .Axes(xlCategory).MinorGridlines.Border.LineStyle = xlDot
      .Axes(xlCategory).MinorGridlines.Border.Weight = xlHairline
      .Axes(xlCategory).MinorGridlines.Border.color = RGB(222, 222, 222)
      
      .Axes(xlCategory).MinimumScale = Range("A33").Value
      .Axes(xlCategory).MaximumScale = Range("A" & endLine).Value + 8
      
      .Axes(xlValue).MinimumScale = 0
      .Axes(xlValue).MaximumScale = WorksheetFunction.Round(Range("B" & endLine).Value + 10, -1)
      

'      GrafCount = GrafCount + 1
'      .SeriesCollection.NewSeries
'      .SeriesCollection(GrafCount).Name = "=""OK数(実際)"""
'      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
'      .SeriesCollection(GrafCount).Values = "='総合結果'!$C$33:$C$" & endLine
'      .SeriesCollection(GrafCount).ChartType = xlXYScatterSmoothNoMarkers
'      .SeriesCollection(GrafCount).Border.Weight = xlThin
'      .SeriesCollection(GrafCount).Border.Color = RGB(0, 0, 255)
'
'       GrafCount = GrafCount + 1
'      .SeriesCollection.NewSeries
'      .SeriesCollection(GrafCount).Name = "=""OK数"""
'      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
'      .SeriesCollection(GrafCount).Values = "='総合結果'!$C$33:$C$" & endLine
'      .SeriesCollection(GrafCount).ChartType = xlXYScatterLinesNoMarkers
'      .SeriesCollection(GrafCount).Border.LineStyle = xlDot
'      .SeriesCollection(GrafCount).Border.Weight = xlThin
'      .SeriesCollection(GrafCount).Border.Color = RGB(0, 0, 150)

      GrafCount = GrafCount + 1
      .SeriesCollection.NewSeries
      .SeriesCollection(GrafCount).Name = "=""NG数"""
      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
      .SeriesCollection(GrafCount).Values = "='総合結果'!$D$33:$D$" & endLine
      .SeriesCollection(GrafCount).ChartType = xlXYScatterSmoothNoMarkers
      .SeriesCollection(GrafCount).Border.Weight = xlThin
      .SeriesCollection(GrafCount).Border.color = RGB(255, 0, 0)
      
      GrafCount = GrafCount + 1
      .SeriesCollection.NewSeries
      .SeriesCollection(GrafCount).Name = "=""NG数(実際)"""
      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
      .SeriesCollection(GrafCount).Values = "='総合結果'!$D$33:$D$" & endLine
      .SeriesCollection(GrafCount).ChartType = xlXYScatterLinesNoMarkers
      .SeriesCollection(GrafCount).Border.LineStyle = xlDot
      .SeriesCollection(GrafCount).Border.Weight = xlThin
      .SeriesCollection(GrafCount).Border.color = RGB(150, 0, 0)

      GrafCount = GrafCount + 1
      .SeriesCollection.NewSeries
      .SeriesCollection(GrafCount).Name = "=""未実施"""
      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
      .SeriesCollection(GrafCount).Values = "='総合結果'!$E$33:$E$" & endLine
      .SeriesCollection(GrafCount).ChartType = xlXYScatterSmoothNoMarkers
      .SeriesCollection(GrafCount).Border.Weight = xlThin
      .SeriesCollection(GrafCount).Border.color = RGB(192, 192, 131)

      GrafCount = GrafCount + 1
      .SeriesCollection.NewSeries
      .SeriesCollection(GrafCount).Name = "=""未実施(実際)"""
      .SeriesCollection(GrafCount).XValues = "='総合結果'!$A$33:$A$" & endLine
      .SeriesCollection(GrafCount).Values = "='総合結果'!$E$33:$E$" & endLine
      .SeriesCollection(GrafCount).ChartType = xlXYScatterLinesNoMarkers
      .SeriesCollection(GrafCount).Border.LineStyle = xlDot
      .SeriesCollection(GrafCount).Border.Weight = xlThin
      .SeriesCollection(GrafCount).Border.color = RGB(0, 0, 150)

  End With

  Worksheets(SetActiveSheet).Select
  Range(SetActiveCell).Select
  If Not BeforeCloseFlg Then
    Worksheets("総合結果").Select
  End If

  
  
  Range("H32").Select
Exit Function

ErrHand:
  
  ' プログレスバーの表示終了処理
  
  ProgressBar_ProgShowClose
  
  If Not BeforeCloseFlg Then
    Worksheets("総合結果").Select
  End If
  
  Call Library_ErrorHandle(Err.Number, Err.Description)

End Function

'***********************************************************************************************************************************************
' * 全シート設定セルを選択状態にする
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_SetCellActive(targetCell As String)

  Dim sheetName As Object
  Dim result As Boolean

  
  ' ======================= 処理開始 ======================
  
  TestSpecification_Init
  
  For Each sheetName In ActiveWorkbook.Sheets

    Worksheets(sheetName.Name).Select
  
    ' テストケース用のシートかどうかチェック
    result = Library_CheckExcludeSheet(sheetName.Name, 9)
    If result = True Then
      Application.Goto Reference:=Range(targetCell), Scroll:=True
    End If
  Next
  
  Worksheets(SetActiveSheet).Select
  

End Function

'***********************************************************************************************************************************************
' * セルのスタイル設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub TestSpecification_SetCellStyles()

  Dim n()
  Dim CheckFlg As Boolean
  Dim ColumLoopArray As Variant
  Dim ColumnLoopName As Variant
  
  On Error GoTo ErrHand
  CheckFlg = False
  
  
  ColumLoopArray = Array("通貨", "パーセント", """青文字", "他システム連動", "注意")
  
  j = ActiveWorkbook.Styles.count
  ReDim n(j)
  For i = 1 To j
    n(i) = ActiveWorkbook.Styles(i).Name
  Next
  For i = 1 To j
    For Each ColumnLoopName In ColumLoopArray
      If n(i) <> ColumnLoopName Then
        ActiveWorkbook.Styles(n(i)).Delete
      Else
        CheckFlg = True
    End If
  Next

  If CheckFlg = False Then
    ActiveWorkbook.Styles.Add Name:="青文字"
    ActiveWorkbook.Styles.Add Name:="他システム連動"
    ActiveWorkbook.Styles.Add Name:="注意"
  End If

  With ActiveWorkbook.Styles("Normal")
    .IncludeNumber = False
    .IncludeFont = True
    .IncludeAlignment = False
    .IncludeBorder = False
    .IncludePatterns = False
    .IncludeProtection = False
        
    .Font.Name = "Meiryo UI"
    .Font.Size = 9
    .Font.Bold = False
    .Font.Italic = False
    .Font.Underline = xlUnderlineStyleNone
    .Font.Strikethrough = False
    .Font.color = RGB(0, 0, 0)
    .Font.TintAndShade = 0
    .Font.ThemeFont = xlThemeFontNone
  End With
  
  
    With ActiveWorkbook.Styles("青文字")
      .IncludeNumber = False
      .IncludeFont = True
      .IncludeAlignment = False
      .IncludeBorder = False
      .IncludePatterns = False
      .IncludeProtection = False
      
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .Font.Bold = False
      .Font.Italic = False
      .Font.Underline = xlUnderlineStyleNone
      .Font.Strikethrough = False
      .Font.color = RGB(0, 0, 255)
      .Font.TintAndShade = 0
      .Font.ThemeFont = xlThemeFontNone
    End With

  
    With ActiveWorkbook.Styles("他システム連動")
      .IncludeNumber = False
      .IncludeFont = True
      .IncludeAlignment = False
      .IncludeBorder = False
      .IncludePatterns = True
      .IncludeProtection = False
      
      .Interior.color = RGB(252, 204, 204)
      .Interior.Pattern = xlSolid
      
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .Font.Bold = False
      .Font.Italic = False
      .Font.Underline = xlUnderlineStyleNone
      .Font.Strikethrough = False
      .Font.color = RGB(0, 0, 0)
      .Font.TintAndShade = 0
      .Font.ThemeFont = xlThemeFontNone
    End With

    With ActiveWorkbook.Styles("注意")
      .IncludeNumber = False
      .IncludeFont = True
      .IncludeAlignment = False
      .IncludeBorder = False
      .IncludePatterns = True
      .IncludeProtection = False
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .Font.Bold = True
      .Font.Italic = False
      .Font.Underline = xlUnderlineStyleNone
      .Font.Strikethrough = False
      .Font.ThemeColor = 1
      .Font.TintAndShade = 0
      .Font.ThemeFont = xlThemeFontNone
      .Interior.Pattern = xlSolid
      .Interior.PatternColorIndex = 0
      .Interior.color = 255
      .Interior.TintAndShade = 0
      .Interior.PatternTintAndShade = 0
    End With
  Exit Sub
  
ErrHand:
  Resume Next
End Sub


'***********************************************************************************************************************************************
' * ステータスバーにテスト結果を表示する
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_DisplayStatusBar()

  Dim MsgString As String
  Dim result As Boolean
  
  Dim SheetCount As Long
  Dim SheetCountOK As Long
  Dim SheetCountNG As Long
  Dim SheetCountNotTest As Long
  Dim SheetCountNotTestStr As String
  
  Application.StatusBar = False
  
  result = Library_CheckExcludeSheet(activeSheet.Name, 8)
  
  
  If result = False Then
    Exit Function
  End If
   Exit Function
  
  SheetCount = Range("F1").Value
  tmp = Split(Range("F2").Value, "/")
  SheetCountNG = val(tmp(0))

  SheetCountOK = Range("F3").Value
  SheetCountNotTestStr = Range("F4").Value
  
  If SheetCount < 0 Then
    Exit Function
  End If
  
  tmp = Split(SheetCountNotTestStr, "/")
  SheetCountNotTest = val(tmp(0))
  
  MsgString = "全件数:" & SheetCount
  MsgString = MsgString & "　NG件数:" & SheetCountNG
  MsgString = MsgString & "　OK件数:" & SheetCountOK
  MsgString = MsgString & "　未実施:" & SheetCountNotTest

  If SheetCountOK > 0 And SheetCountNotTest > 0 Then
    Application.StatusBar = MsgString
  ElseIf SheetCountNotTest = 0 Then
    Application.StatusBar = "テスト完了"
  End If
End Function


'***********************************************************************************************************************************************
' * 高さの調節
' *
' * @param  Long  RowCnt     ：列番号
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function TestSpecification_setCellHeight(RowCnt As Long)
    '
  Rows(RowCnt & ":" & RowCnt).EntireRow.AutoFit
  If Rows(RowCnt & ":" & RowCnt).Height < SetRowHeight Then
    Rows(RowCnt & ":" & RowCnt).RowHeight = SetRowHeight + SetAddRowHeight
  Else
    Rows(RowCnt & ":" & RowCnt).RowHeight = Rows(RowCnt & ":" & RowCnt).Height + SetAddRowHeight
  End If
  If Len(Range("D" & RowCnt).Value) = 58 Or Len(Range("E" & RowCnt).Value) = 58 Then
    Rows(RowCnt & ":" & RowCnt).RowHeight = Rows(RowCnt & ":" & RowCnt).Height + SetRowHeight + SetAddRowHeight
  End If
End Function

