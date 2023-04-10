Attribute VB_Name = "Ctl_Style"
Option Explicit

Dim setStyleBook     As Workbook


'**************************************************************************************************
' * スタイルImport/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function スタイル出力()
  Dim FilePath As String, fileName As String
  Const funcName As String = "Ctl_Style.スタイル出力"
     
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  LadexSh_Style.copy
  
  Set setStyleBook = ActiveWorkbook
  setStyleBook.SaveAs LadexDir & "\" & "スタイル情報.xlsx"
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", FilePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時-------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function スタイル取込()
  Dim styleBookPath As String
  Dim FilePath As String, fileName As String
  Const funcName As String = "Ctl_Style.スタイル取込"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Ctl_ProgressBar.showStart
  PrgP_Max = 4
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  If Library.chkIsOpen("スタイル情報.xlsx") Then
    Set setStyleBook = Workbooks("スタイル情報.xlsx")
    setStyleBook.Save
  Else
    Set setStyleBook = Workbooks.Open(LadexDir & "\" & "スタイル情報.xlsx")
    Call Library.startScript
  End If
  setStyleBook.Sheets("Style").Columns("A:J").copy LadexBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  styleBookPath = setStyleBook.path & "\" & setStyleBook.Name
  Application.DisplayAlerts = False
  setStyleBook.Close
'  Call Library.execDel(styleBookPath)
  
  Set setStyleBook = Nothing
  If MsgBox("スタイルを適応しますか？", vbYesNo + vbExclamation) = vbYes Then
    Call Ctl_Style.スタイル設定
  End If
  
  
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
'エラー発生時-------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'**************************************************************************************************
' * スタイル削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function スタイル削除()
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  Dim useStyleName As Variant
  
  Const funcName As String = "Ctl_Style.スタイル削除"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("スタイル利用確認")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  'ブックの保護確認
  If ActiveWorkbook.ProtectWindows = True Then
    Call Library.showNotice(412, , True)
  End If

  'シートの保護確認
  For Each tempSheet In Sheets
    If Worksheets(tempSheet.Name).ProtectContents = True Then
      Worksheets(tempSheet.Name).Select
      Call Library.showNotice(413, , True)
    End If
  Next
  
  Call Library.スタイル利用確認
  Call Library.スタイル削除

  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function スタイル_全削除()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  Dim useStyleName As Variant
  
  Const funcName As String = "Ctl_Style.スタイル_全削除"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("スタイル利用確認")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  'ブックの保護確認
  If ActiveWorkbook.ProtectWindows = True Then
    Call Library.showNotice(412, , True)
  End If

  'シートの保護確認
  For Each tempSheet In Sheets
    If Worksheets(tempSheet.Name).ProtectContents = True Then
      Worksheets(tempSheet.Name).Select
      Call Library.showNotice(413, , True)
    End If
  Next
  
  count = 1
  endCount = ActiveWorkbook.Styles.count
  
    
  ReDim useStyle(0)
  useStyle(0) = "標準"
  Call Library.スタイル削除

  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'**************************************************************************************************
' * スタイル設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function スタイル設定()
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
  Const funcName As String = "Ctl_Style.スタイル設定"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("スタイル利用確認")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  '既存スタイル削除------------------------------
  Call Library.スタイル利用確認
  Call Library.スタイル削除
  
  If setStyleBook Is Nothing Then
    If Library.chkIsOpen("スタイル情報.xlsx") Then
      Set setStyleBook = Workbooks("スタイル情報.xlsx")
      setStyleBook.Save
    Else
      Set setStyleBook = Workbooks.Open(LadexDir & "\" & "スタイル情報.xlsx")
      Call Library.startScript
    End If
    setStyleBook.Sheets("Style").Columns("A:J").copy LadexBook.Worksheets("Style").Range("A1")
    setStyleBook.Close
  End If
  Set setStyleBook = Nothing
  
  
  'スタイル初期化--------------------------------
  endLine = LadexSh_Style.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    On Error Resume Next
    
    If LadexSh_Style.Range("A" & line) <> "無効" Then
      Call Ctl_ProgressBar.showCount("スタイル設定", 1, 2, line, endLine, LadexSh_Style.Range("B" & line))

      Select Case LadexSh_Style.Range("B" & line)
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
          Call Library.showDebugForm("定義済スタイル", LadexSh_Style.Range("B" & line), "debug")
          
      'Ladexの初期設定
      Case "桁区切り", "パーセント", "通貨", "通貨[千単位]", "数値", "数値[千単位]", "00.0", "日付 [yyyy/mm/dd]", "日付 [yyyy/m]", "日時", "不要", "Error", "要確認", "H_標準", "H_目次1", "H_目次2", "H_目次3", "《》"
        Call Library.showDebugForm("Ladexスタイル ", LadexSh_Style.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=LadexSh_Style.Range("B" & line).Value
      Case Else
        Call Library.showDebugForm("スタイル名", LadexSh_Style.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=LadexSh_Style.Range("B" & line).Value
      End Select

      With ActiveWorkbook.Styles(LadexSh_Style.Range("B" & line).Value)

        If LadexSh_Style.Range("C" & line) <> "" Then
          .NumberFormatLocal = LadexSh_Style.Range("C" & line)
        End If

        .IncludeNumber = LadexSh_Style.Range("D" & line)
        .IncludeFont = LadexSh_Style.Range("E" & line)
        .IncludeAlignment = LadexSh_Style.Range("F" & line)
        .IncludeBorder = LadexSh_Style.Range("G" & line)
        .IncludePatterns = LadexSh_Style.Range("H" & line)
        .IncludeProtection = LadexSh_Style.Range("I" & line)

        If LadexSh_Style.Range("E" & line) = "TRUE" Then
          .Font.Name = LadexSh_Style.Range("J" & line).Font.Name
          .Font.Size = LadexSh_Style.Range("J" & line).Font.Size
          .Font.Color = LadexSh_Style.Range("J" & line).Font.Color
          .Font.Bold = LadexSh_Style.Range("J" & line).Font.Bold
        End If

        '配置
        If LadexSh_Style.Range("F" & line) = "TRUE" Then
          .HorizontalAlignment = LadexSh_Style.Range("J" & line).HorizontalAlignment
          .VerticalAlignment = LadexSh_Style.Range("J" & line).VerticalAlignment
        End If

        '罫線
        If LadexSh_Style.Range("G" & line) = "TRUE" Then
          If LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).LineStyle <> xlNone Then
            .Borders(xlDiagonalDown).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).LineStyle
            .Borders(xlDiagonalDown).Weight = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).Weight
            .Borders(xlDiagonalDown).Color = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).LineStyle <> xlNone Then
            .Borders(xlDiagonalUp).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).LineStyle
            .Borders(xlDiagonalUp).Weight = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).Weight
            .Borders(xlDiagonalUp).Color = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlLeft).LineStyle <> xlNone Then
            .Borders(xlLeft).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlLeft).LineStyle
            .Borders(xlLeft).Weight = LadexSh_Style.Range("J" & line).Borders(xlLeft).Weight
            .Borders(xlLeft).Color = LadexSh_Style.Range("J" & line).Borders(xlLeft).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlRight).LineStyle <> xlNone Then
            .Borders(xlRight).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlRight).LineStyle
            .Borders(xlRight).Weight = LadexSh_Style.Range("J" & line).Borders(xlRight).Weight
            .Borders(xlRight).Color = LadexSh_Style.Range("J" & line).Borders(xlRight).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlTop).LineStyle <> xlNone Then
            .Borders(xlTop).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlTop).LineStyle
            .Borders(xlTop).Weight = LadexSh_Style.Range("J" & line).Borders(xlTop).Weight
            .Borders(xlTop).Color = LadexSh_Style.Range("J" & line).Borders(xlTop).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlBottom).LineStyle <> xlNone Then
            .Borders(xlBottom).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlBottom).LineStyle
            .Borders(xlBottom).Weight = LadexSh_Style.Range("J" & line).Borders(xlBottom).Weight
            .Borders(xlBottom).Color = LadexSh_Style.Range("J" & line).Borders(xlBottom).Color
          End If
        End If


        '背景色
        If LadexSh_Style.Range("H" & line) = "TRUE" Then
          .Interior.Color = LadexSh_Style.Range("J" & line).Interior.Color
        End If
      End With
    End If
  Next
  On Error GoTo catchError
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function スタイル初期化()
  Dim FSO As Object
  Dim setActivBook     As Workbook
  Dim FilePath As String, fileName As String
  
  Const funcName As String = "Ctl_Style.スタイル初期化"

  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("スタイル利用確認")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.スタイル削除

  Set setActivBook = ActiveWorkbook
  Set setStyleBook = Workbooks.add
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  '新規ブックを作成しスタイルをインポート--------
  With setStyleBook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      FilePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs FilePath
  End With
  Call Library.showDebugForm("新規ブック作成", FilePath, "debug")
  Call Library.showDebugForm("スタイルのマージ", , "debug")
  
  
  setActivBook.Activate
  ActiveWorkbook.Styles.Merge Workbook:=Workbooks(fileName)
  Set FSO = Nothing
  setStyleBook.Close
  
  Call Library.execDel(FilePath)
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function 標準スタイルの見た目変更()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Style.標準スタイルの見た目変更"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  With ActiveWorkbook.Styles("Percent")
    .IncludeNumber = True
    .IncludeFont = True
    .IncludeAlignment = True
    .IncludeBorder = True
    .IncludePatterns = True
    .IncludeProtection = True
  
    'フォント
    .Font.Name = "メイリオ"
    .Font.Size = 9
    .Font.Color = ""
    .Font.Bold = ""
  
    '配置
    .HorizontalAlignment = ""
    .VerticalAlignment = ""
  
    '罫線
    .Borders(xlDiagonalDown).LineStyle = ""
    .Borders(xlDiagonalDown).Weight = ""
    .Borders(xlDiagonalDown).Color = ""
    .Borders(xlDiagonalUp).LineStyle = ""
    .Borders(xlDiagonalUp).Weight = ""
    .Borders(xlDiagonalUp).Color = ""
  
    .Borders(xlLeft).LineStyle = ""
    .Borders(xlLeft).Weight = ""
    .Borders(xlLeft).Color = ""
  
    .Borders(xlRight).LineStyle = ""
    .Borders(xlRight).Weight = ""
    .Borders(xlRight).Color = ""
  
    .Borders(xlTop).LineStyle = ""
    .Borders(xlTop).Weight = ""
    .Borders(xlTop).Color = ""
  
    .Borders(xlBottom).LineStyle = ""
    .Borders(xlBottom).Weight = ""
    .Borders(xlBottom).Color = ""
  
    '背景色
    .Interior.Color = ""
   End With
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function スタイル確認()
  Dim styleCnt As Long
  Dim objSheet As Variant
  Dim objStyle As Variant
  Dim sheetName As String, styleName As String
  Dim slctRange As Range
  Dim RangeCnt As Long, RangeAllCnt As Long
  Dim chkNewSheetFlg As Boolean
  
  
  Const funcName As String = "Ctl_Style.スタイル確認"
  Const AbortCnt As Long = 10000
     
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("スタイル利用確認")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  styleCnt = 1
  chkNewSheetFlg = False
  Set useStyleVal = Nothing
  Set useStyleVal = CreateObject("Scripting.Dictionary")
  
  '利用スタイルの取得---------------------------
  Call Library.スタイル利用確認
  
  For Each objStyle In ActiveWorkbook.Styles
    styleName = objStyle.Name
    Call Library.showDebugForm("スタイル名", styleName, "debug")
    
    If Library.chkArrayVal(useStyle(), styleName) = True Then
      Select Case styleName
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Case Else
          For Each objSheet In ActiveWorkbook.Sheets
            sheetName = objSheet.Name
            If Worksheets(sheetName).Visible = xlSheetVisible Then
              Worksheets(sheetName).Select
              Cells(Rows.count, Columns.count).Select
            
              chkNewSheetFlg = True
              RangeCnt = 1
              RangeAllCnt = Worksheets(sheetName).UsedRange.count
              
              For Each slctRange In Worksheets(sheetName).UsedRange
                If styleName = slctRange.style Then
                  If chkNewSheetFlg = True Then
                    slctRange.Select
                    chkNewSheetFlg = False
                  Else
                    Application.Union(Selection, slctRange).Select
                  End If
                End If
                'Call Library.showDebugForm("利用スタイル", slctRange.style, "debug")
                Call Ctl_ProgressBar.showBar("スタイル確認", styleCnt, ActiveWorkbook.Styles.count, RangeCnt, RangeAllCnt, styleName)
                If RangeCnt >= AbortCnt Then
                  Exit For
                End If
                RangeCnt = RangeCnt + 1
              Next
              
              If Selection.Address <> Cells(Rows.count, Columns.count).Address Then
                If useStyleVal.Exists(styleName) Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>" & sheetName & "!" & Selection.Address
                Else
                  useStyleVal.add styleName, sheetName & "!" & Selection.Address
                End If
                If RangeCnt >= AbortCnt Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>Abort" & "!" & sheetName & "シートでの設定数が多数存在するため、調査を中断"
                End If
              ElseIf Selection.Address = Cells(Rows.count, Columns.count).Address Then
              Else
                If useStyleVal.Exists(styleName) Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>Abort" & "!" & sheetName & "シートでの設定数が多数存在するため、調査を中断"
                Else
                  useStyleVal.add styleName, "Abort" & "!" & sheetName & "シートでの設定数が多数存在するため、調査を中断"
                End If
              End If
            End If
          Next
      End Select
      styleCnt = styleCnt + 1
    End If
  Next
  
  Call Ctl_ProgressBar.showEnd
  If useStyleVal.count > 0 Then
    Frm_Style.Show vbModeless
  Else
    Call Library.showNotice(10, "利用されているスタイルが見つかりませんでした")
  End If
  
  Call Ctl_Book.A1セル選択

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
'    Call init.resetGlobalVal
  End If
  Exit Function
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
