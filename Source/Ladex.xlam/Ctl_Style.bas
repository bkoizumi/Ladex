Attribute VB_Name = "Ctl_Style"
Option Explicit

Dim setStyleBook     As Workbook


'**************************************************************************************************
' * スタイルImport/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Export()
  Dim filePath As String, fileName As String
  Const funcName As String = "Ctl_Style.Export"
     
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
  BK_sheetStyle.Copy
  
  Set setStyleBook = ActiveWorkbook
  setStyleBook.SaveAs LadexDir & "\" & "スタイル情報.xlsx"
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", filePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function Import()
  Dim styleBookPath As String
  Dim filePath As String, fileName As String
  Const funcName As String = "Ctl_Style.Import"
  
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
  setStyleBook.Sheets("Style").Columns("A:J").Copy BK_ThisBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  styleBookPath = setStyleBook.Path & "\" & setStyleBook.Name
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
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'**************************************************************************************************
' * スタイル削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function スタイル削除()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  Const funcName As String = "Ctl_Style.スタイル削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
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
  
  For Each s In ActiveWorkbook.Styles
    Call Ctl_ProgressBar.showCount("定義済スタイル削除", 1, 2, count, endCount, s.Name)
    Select Case s.Name
      Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Call Library.showDebugForm("定義済スタイル", s.Name, "debug")
      
      'Ladexの初期設定
      Case "桁区切り", "パーセント", "通貨", "通貨[千単位]", "数値", "数値[千単位]", "00.0", "日付 [yyyy/mm/dd]", "日付 [yyyy/m]", "日時", "不要", "Error", "要確認", "H_標準", "H_目次1", "H_目次2", "H_目次3", "《》"
        Call Library.showDebugForm("Ladexスタイル ", s.Name, "debug")
      
      Case Else
        Call Library.showDebugForm("削除スタイル  ", s.Name, "debug")
        s.delete
    End Select
    count = count + 1
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'**************************************************************************************************
' * スタイル設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function スタイル設定()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
'  On Error Resume Next
  Const funcName As String = "Ctl_Style.スタイル設定"
  
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call Ctl_Style.スタイル削除
  
  'スタイル初期化--------------------------------
  endLine = BK_sheetStyle.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetStyle.Range("A" & line) <> "無効" Then
      Call Ctl_ProgressBar.showCount("スタイル設定", 1, 2, line, endLine, BK_sheetStyle.Range("B" & line))

      Select Case BK_sheetStyle.Range("B" & line)
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
          Call Library.showDebugForm("定義済スタイル", BK_sheetStyle.Range("B" & line), "debug")
          
      'Ladexの初期設定
      Case "桁区切り", "パーセント", "通貨", "通貨[千単位]", "数値", "数値[千単位]", "00.0", "日付 [yyyy/mm/dd]", "日付 [yyyy/m]", "日時", "不要", "Error", "要確認", "H_標準", "H_目次1", "H_目次2", "H_目次3", "《》"
        Call Library.showDebugForm("Ladexスタイル ", BK_sheetStyle.Range("B" & line), "debug")
      
      Case Else
        Call Library.showDebugForm("スタイル名", BK_sheetStyle.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=BK_sheetStyle.Range("B" & line).Value
      End Select

      With ActiveWorkbook.Styles(BK_sheetStyle.Range("B" & line).Value)

        If BK_sheetStyle.Range("C" & line) <> "" Then
          .NumberFormatLocal = BK_sheetStyle.Range("C" & line)
        End If

        .IncludeNumber = BK_sheetStyle.Range("D" & line)
        .IncludeFont = BK_sheetStyle.Range("E" & line)
        .IncludeAlignment = BK_sheetStyle.Range("F" & line)
        .IncludeBorder = BK_sheetStyle.Range("G" & line)
        .IncludePatterns = BK_sheetStyle.Range("H" & line)
        .IncludeProtection = BK_sheetStyle.Range("I" & line)

        If BK_sheetStyle.Range("E" & line) = "TRUE" Then
          .Font.Name = BK_sheetStyle.Range("J" & line).Font.Name
          .Font.Size = BK_sheetStyle.Range("J" & line).Font.Size
          .Font.Color = BK_sheetStyle.Range("J" & line).Font.Color
          .Font.Bold = BK_sheetStyle.Range("J" & line).Font.Bold
        End If

        '配置
        If BK_sheetStyle.Range("F" & line) = "TRUE" Then
          .HorizontalAlignment = BK_sheetStyle.Range("J" & line).HorizontalAlignment
          .VerticalAlignment = BK_sheetStyle.Range("J" & line).VerticalAlignment
        End If

        '罫線
        If BK_sheetStyle.Range("G" & line) = "TRUE" Then
          If BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).LineStyle <> xlNone Then
            .Borders(xlDiagonalDown).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).LineStyle
            .Borders(xlDiagonalDown).Weight = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).Weight
            .Borders(xlDiagonalDown).Color = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).Color
          End If

          If BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).LineStyle <> xlNone Then
            .Borders(xlDiagonalUp).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).LineStyle
            .Borders(xlDiagonalUp).Weight = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).Weight
            .Borders(xlDiagonalUp).Color = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).Color
          End If

          If BK_sheetStyle.Range("J" & line).Borders(xlLeft).LineStyle <> xlNone Then
            .Borders(xlLeft).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlLeft).LineStyle
            .Borders(xlLeft).Weight = BK_sheetStyle.Range("J" & line).Borders(xlLeft).Weight
            .Borders(xlLeft).Color = BK_sheetStyle.Range("J" & line).Borders(xlLeft).Color
          End If

          If BK_sheetStyle.Range("J" & line).Borders(xlRight).LineStyle <> xlNone Then
            .Borders(xlRight).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlRight).LineStyle
            .Borders(xlRight).Weight = BK_sheetStyle.Range("J" & line).Borders(xlRight).Weight
            .Borders(xlRight).Color = BK_sheetStyle.Range("J" & line).Borders(xlRight).Color
          End If

          If BK_sheetStyle.Range("J" & line).Borders(xlTop).LineStyle <> xlNone Then
            .Borders(xlTop).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlTop).LineStyle
            .Borders(xlTop).Weight = BK_sheetStyle.Range("J" & line).Borders(xlTop).Weight
            .Borders(xlTop).Color = BK_sheetStyle.Range("J" & line).Borders(xlTop).Color
          End If

          If BK_sheetStyle.Range("J" & line).Borders(xlBottom).LineStyle <> xlNone Then
            .Borders(xlBottom).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlBottom).LineStyle
            .Borders(xlBottom).Weight = BK_sheetStyle.Range("J" & line).Borders(xlBottom).Weight
            .Borders(xlBottom).Color = BK_sheetStyle.Range("J" & line).Borders(xlBottom).Color
          End If
        End If


        '背景色
        If BK_sheetStyle.Range("H" & line) = "TRUE" Then
          .Interior.Color = BK_sheetStyle.Range("J" & line).Interior.Color
        End If
      End With
    End If
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function スタイル初期化()
  Dim FSO As Object
  Dim setActivBook     As Workbook
  Dim filePath As String, fileName As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Style.スタイル初期化"

  Call Library.startScript
  Call init.setting
  Call Library.showDebugForm(funcName & "開始==========================================")
  '----------------------------------------------
  Call Ctl_Style.スタイル削除

  Set setActivBook = ActiveWorkbook
  Set setStyleBook = Workbooks.add
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  With setStyleBook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      filePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs filePath
  End With
  
  setActivBook.Activate
  ActiveWorkbook.Styles.Merge Workbook:=Workbooks(fileName)
  Set FSO = Nothing
  setStyleBook.Close
  
  Call Library.execDel(filePath)
  
  '処理終了--------------------------------------
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("", , "end")
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function
