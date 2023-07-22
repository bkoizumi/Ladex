Attribute VB_Name = "Ctl_Cells"
Option Explicit

'**************************************************************************************************
' * セル調整
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function セル自動調整_幅()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  Dim colName As String
  
  Const funcName As String = "Ctl_Cells.セル自動調整_幅"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  Columns(Library.getColumnName(startColLine) & ":" & Library.getColumnName(endColLine)).EntireColumn.AutoFit
  
  '最大値確認------------------------------------
  For colLine = startLine To endColLine
    colName = Library.getColumnName(colLine)
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, colLine, endColLine, "")
    
    If Columns(colLine).Hidden = False Then
      '列幅の自動調整
      Columns(colLine).EntireColumn.AutoFit
    
      If Columns(colLine).ColumnWidth > maxColumnWidth Then
        Columns(colLine).ColumnWidth = maxColumnWidth
      End If
    End If
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル自動調整_高さ()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル自動調整_高さ"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
    
    If Rows(line).Hidden = False Then
      '高さの自動調整
      Rows(line).EntireRow.AutoFit
      
      If Rows(line).Height > maxRowHeight Then
        Rows(line).rowHeight = maxRowHeight
      End If
    End If
  Next
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル自動調整_両方()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル自動調整_両方"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Max = 2
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Ctl_Cells.セル自動調整_高さ
  Call Ctl_Cells.セル自動調整_幅
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル固定設定_幅()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル固定設定_幅"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  Columns(Library.getColumnName(startColLine) & ":" & Library.getColumnName(endColLine)).EntireColumn.AutoFit
  
  '最大値確認------------------------------------
  For colLine = startLine To endColLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, colLine, endColLine, "")
    
    If Columns(colLine).Hidden = False Then
      Columns(colLine).ColumnWidth = dicVal("ColumnWidth")
    End If
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル固定設定_高さ()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル固定設定_高さ"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.getCellSelectArea(startLine, endLine, startColLine, endColLine)
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
    If Rows(line).Hidden = False Then
      Rows(line).rowHeight = dicVal("rowHeight")
    End If
  Next
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル固定設定_両方()
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル固定設定_両方"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  PrgP_Max = 2
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  Ctl_Cells.セル固定設定_高さ
  Ctl_Cells.セル固定設定_幅
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function セル固定設定_高さ指定(heightVal As Integer)
  Dim startLine As Long, endLine As Long, startColLine As Long, endColLine As Long
  Dim line As Long, colLine As Long
  
  Const funcName As String = "Ctl_Cells.セル固定設定_高さ"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")

  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  '選択範囲がある場合----------------------------
  If Selection.Rows.count > 1 Then
    startLine = Selection(1).Row
    endLine = Selection(Selection.count).Row
  
    For line = startLine To endLine
      Call Ctl_ProgressBar.showBar(funcName, PrgP_Cnt, PrgP_Max, line, endLine, "")
      
      If Rows(line).Hidden = False Then
        Rows(line).rowHeight = heightVal
      End If
    Next
  Else
    Cells.rowHeight = heightVal
  End If
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
' * セル編集
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 削除_前後のスペース()
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.削除_前後のスペース"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    slctCells.Value = Trim(slctCells.Text)
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 削除_全スペース()
  Dim slctCells As Range
  Dim resVal As String
  
  Const funcName As String = "Ctl_Cells.削除_全スペース"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    resVal = slctCells.Text

    If resVal <> "" Then
      resVal = Replace(resVal, " ", "")
      resVal = Replace(resVal, "　", "")
      slctCells.Value = resVal
      DoEvents
    End If
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 追加_文頭に中黒点()
  Dim line As Long, endLine As Long
  Dim Reg As Object
  Dim slctCells
  
  Const funcName As String = "Ctl_Cells.追加_文頭に中黒点"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^・"
    .IgnoreCase = False
    .Global = True
  End With
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    slctCells.Value = "・" & Reg.Replace(slctCells.Value, "")
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 追加_文頭に連番()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.追加_文頭に連番"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^[0-9]+．"
    .IgnoreCase = False
    .Global = True
  End With

  line = 1
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    Call Library.showDebugForm("設定前セル値", slctCells.Value, "debug")
    
    slctCells.NumberFormatLocal = "@"
    slctCells.Value = line & "．" & Reg.Replace(slctCells.Value, "")
    line = line + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 上書_文頭に連番()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim i As Long
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.上書_文頭に連番"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  line = 1
  
  If Selection.Item(1).Value = "" Then
    line = 1
  Else
    line = Selection.Item(1).Value
  End If
  
  Selection.HorizontalAlignment = xlCenter
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    Call Library.showDebugForm("設定前セル値", slctCells.Value, "debug")
    slctCells.Value = line
    line = line + 1
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_全角⇒半角()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.変換_全角⇒半角"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = Library.convZen2Han(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_半角⇒全角()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.変換_半角⇒全角"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = Library.convHan2Zen(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 設定_取り消し線()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.設定_取り消し線"
    
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)

    If slctCells.Font.Strikethrough = True Then
      slctCells.Font.Strikethrough = False
    Else
      slctCells.Font.Strikethrough = True
    End If
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function コメント_追加()
  Dim commentVal As String, commentBgColor As Long, CommentFontColor As Long
  Dim CommentFont As String, CommentFontSize As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント_追加"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  For Each slctCells In Selection
    
    '結合されている場合
    If slctCells.MergeCells Then
      If slctCells.MergeArea.Item(1).Address = slctCells.Address Then
      Else
        GoTo LBl_nextFor
      End If
    End If

    If TypeName(slctCells.Comment) = "Comment" Then
      commentVal = slctCells.Comment.Text
      commentBgColor = slctCells.Comment.Shape.Fill.ForeColor.RGB
      CommentFontSize = slctCells.Comment.Shape.TextFrame.Characters.Font.Size
      CommentFont = slctCells.Comment.Shape.TextFrame.Characters.Font.Name
      CommentFontColor = slctCells.Comment.Shape.TextFrame.Characters.Font.Color
    End If
    
    If FrmVal("commentVal") <> "" Then
      commentVal = FrmVal("commentVal")
    End If
      
    With Frm_InsComment
      .TextBox = commentVal
      If commentVal <> "" Then
        .CommentColor.BackColor = commentBgColor
        .CommentFont = CommentFont
        .CommentFontColor.BackColor = CommentFontColor
        .CommentFontSize = CommentFontSize
      End If
      .Label1.Caption = "選択セル：" & slctCells.Address(RowAbsolute:=False, ColumnAbsolute:=False)
      .Show
    End With
    DoEvents

LBl_nextFor:
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function コメント_画像追加()
  Dim insImgPath As String
  
  Const funcName As String = "Ctl_Cells.コメント_画像追加"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  insImgPath = Library.getFilePath(ActiveWorkbook.path, "", "コメントに画像挿入", "img")
  If insImgPath = "" Then
    GoTo LB_ExitFunction
  End If

  With ActiveCell
    If TypeName(.Comment) = "Comment" Then
      .ClearComments
    End If

    With .AddComment
      .Shape.Fill.UserPicture insImgPath
'      .Shape.Height = 500
'      .Shape.Width = 500
    End With
  End With


  
  '処理終了--------------------------------------
LB_ExitFunction:
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function コメント_削除()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント_削除"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  If ActiveSheet.ProtectContents = True Then
    Call Library.showNotice(413, , True)
  End If
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, "")

    If TypeName(slctCells.Comment) = "Comment" Then
      slctCells.ClearComments
    End If
    DoEvents
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 貼付_行例入れ替え()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.貼付_行例入れ替え"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 上書_ゼロ()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.上書_ゼロ"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    If slctCells.Text = "" Then
      slctCells.Value = 0
      DoEvents
    End If
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 削除_改行()
  Dim slctCells As Range
  Dim resVal As String
  Dim i As Long
  
  Const funcName As String = "Ctl_Cells.削除_改行"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  arrCells = Selection.Value
  For i = LBound(arrCells, 1) To UBound(arrCells, 1)
    arrCells(i, 1) = Replace(arrCells(i, 1), vbCrLf, "")
    arrCells(i, 1) = Replace(arrCells(i, 1), vbCr, "")
    arrCells(i, 1) = Replace(arrCells(i, 1), vbLf, "")
    
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, UBound(arrCells, 1), "")
  Next
  Selection.Value = arrCells
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 挿入_行()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.挿入_行"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 挿入_列()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.挿入_列"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 削除_定数()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.削除_定数"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  On Error Resume Next
  If Selection.CountLarge = 1 Then
    Call Library.showNotice(600)
  ElseIf Selection.CountLarge > 1 Then
    Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents
  End If
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_大文字⇒小文字()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.変換_大文字⇒小文字"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
 
  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = LCase(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_小文字⇒大文字()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.変換_小文字⇒大文字"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0

  For Each arrCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, arrCells.Text)

    If arrCells.Value <> "" Then
      arrCells.Value = UCase(arrCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_丸数字⇒数値()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.変換_丸数字⇒数値"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  slctCellsCnt = 0
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    If slctCells.Value <> "" Then
      Call Library.showDebugForm("Value", AscW(slctCells.Value), "debug")
      Select Case AscW(slctCells.Value)
        Case 9450
          slctCells.Value = 0
        
        '1～20
        Case 9312 To 9332
          slctCells.Value = AscW(slctCells.Value) - 9311
        
        '21～35
        Case 12881 To 12901
          slctCells.Value = AscW(slctCells.Value) - 12881 + 21
  
        '36～50
        Case 12977 To 13027
          slctCells.Value = AscW(slctCells.Value) - 12941
          
        Case Else
      End Select
    End If
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_数値⇒丸数字()
  Dim line As Long, endLine As Long
  Dim slctCells As Range

  Const funcName As String = "Ctl_Cells.変換_数値⇒丸数字"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    Select Case slctCells.Value
      Case 1 To 20
        slctCells.Value = Chr(Asc("①") + slctCells.Value - 1)
        
      Case 21 To 35
        slctCells.Value = ChrW(12881 + slctCells.Value - 21)
      
      Case 36 To 50
        slctCells.Value = ChrW(12941 + slctCells.Value)
        
      Case 0
        slctCells.Value = ChrW(9450)
      Case Else
    End Select
  
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function 変換_URLエンコード()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.変換_URLエンコード"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    slctCells.Value = Library.convURLEncode(slctCells.Text)
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_URLデコード()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.変換_URLデコード"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    slctCells.Value = Library.convURLDecode(slctCells.Text)
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_Unicodeエスケープ()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.変換_Unicodeエスケープ"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)
    
    slctCells.Value = Library.convUnicodeEscape(slctCells.Text)
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function 変換_Unicodeアンエスケープ()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.変換_Unicodeアンエスケープ"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showCount(funcName, PrgP_Cnt, PrgP_Max, PbarCnt, Selection.count, slctCells.Text)


    slctCells.Value = Library.convUnicodeunEscape(slctCells.Text)
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
'@link   https://calmdays.net/vbatips/base64encode/
Function 変換_Base64エンコード()
  Dim slctCells As Range
  Dim oXmlNode  As Object
  
  Const funcName As String = "Ctl_Cells.変換_Base64エンコード"
  
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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  For Each slctCells In Selection
    Set oXmlNode = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    oXmlNode.dataType = "bin.base64"
    oXmlNode.nodeTypedValue = convStringToBinary(slctCells.Value)
    slctCells.Value = Replace(oXmlNode.Text, "77u/", "")
    Set oXmlNode = Nothing
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
'@link   https://calmdays.net/vbatips/base64decode/
Function 変換_Base64デコード()
  Dim slctCells As Range
  Dim oXmlNode  As Object
  
  Const funcName As String = "Ctl_Cells.変換_Base64デコード"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  For Each slctCells In Selection
    Set oXmlNode = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    
    oXmlNode.dataType = "bin.base64"
    oXmlNode.Text = slctCells.Value
    slctCells.Value = convBinaryToString(oXmlNode.nodeTypedValue)
    
    Set oXmlNode = Nothing
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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


'******************************************************************************
'　バイナリを文字列に変換
'@link   https://calmdays.net/vbatips/base64decode/
'******************************************************************************
Function convBinaryToString(getBinary As Variant) As Variant
  Dim oAdoStm As Object
  Set oAdoStm = CreateObject("ADODB.Stream")

  oAdoStm.Type = 1
  oAdoStm.Open
  oAdoStm.Write getBinary
  oAdoStm.Position = 0
  oAdoStm.Type = 2
  oAdoStm.Charset = "utf-8" 'or "shift_jis"
  convBinaryToString = oAdoStm.ReadText
  
  Set oAdoStm = Nothing
End Function


'******************************************************************************
'　文字列をバイナリに変換
'@link   https://calmdays.net/vbatips/base64encode/
'******************************************************************************
Function convStringToBinary(getText As String) As Variant
  Dim oAdoStm As Object
  Set oAdoStm = CreateObject("ADODB.Stream")
  
  oAdoStm.Type = 2
  oAdoStm.Charset = "utf-8" 'or "shift_jis"
  oAdoStm.Open
  oAdoStm.WriteText getText
  oAdoStm.Position = 0
  oAdoStm.Type = 1
  convStringToBinary = oAdoStm.Read
  
  Set oAdoStm = Nothing
End Function
