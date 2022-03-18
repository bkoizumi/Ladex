Attribute VB_Name = "Ctl_Cells"
Option Explicit

'**************************************************************************************************
' * 文字列操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Trim01()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.Trim01"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctCells In Selection
    slctCells.Value = Trim(slctCells.Text)
    DoEvents
  Next

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 全空白削除()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.Trim01"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctCells In Selection
    resVal = slctCells.Text
    
    If resVal <> "" Then
      resVal = Replace(resVal, " ", "")
      resVal = Replace(resVal, "　", "")
      slctCells.Value = resVal
      DoEvents
    End If
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 中黒点付与()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.中黒点付与"
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^・"
    .IgnoreCase = False
    .Global = True
  End With
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each slctCells In Selection
    slctCells.Value = "・" & Reg.Replace(slctCells.Value, "")
    DoEvents
  Next

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function 連番設定()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.連番設定"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  line = 1
  For Each slctCells In Selection
    slctCells.NumberFormatLocal = "###"
    slctCells.Value = line
    line = line + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function 連番追加()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.連番追加"
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^[0-9]+．"
    .IgnoreCase = False
    .Global = True
  End With
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
    slctCells.NumberFormatLocal = "@"
    slctCells.Value = line & "．" & Reg.Replace(slctCells.Value, "")
    line = line + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function 英数字全半角変換()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.英数字全半角変換"

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
  slctCellsCnt = 0
  
  For Each slctCells In Selection
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, slctCellsCnt, Selection.count, "英数字全半角変換")
    If slctCells.Value <> "" Then
      slctCells.Value = Library.convHan2Zen(slctCells.Value)
    End If
    slctCellsCnt = slctCellsCnt + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function 取り消し線設定()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.取り消し線設定"
    
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  For Each slctCells In Selection
    If slctCells.Font.Strikethrough = True Then
      slctCells.Font.Strikethrough = False
    Else
      slctCells.Font.Strikethrough = True
    End If
    DoEvents
  Next

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function コメント挿入()
  Dim commentVal As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント挿入"

  '処理終了--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.startScript
    Call Library.showDebugForm("", , "start")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "start1")
  End If
  '----------------------------------------------
  For Each slctCells In Selection
    commentVal = ""
    If TypeName(slctCells.Comment) = "Comment" Then
      commentVal = slctCells.Comment.Text
    End If
    With Frm_InsComment
      .TextBox = commentVal
      .Label1.Caption = "選択セル：" & slctCells.Address(RowAbsolute:=False, ColumnAbsolute:=False)
      .Show
    End With
    DoEvents
  Next
  
  '処理終了--------------------------------------
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
  Call Library.errorHandle
End Function

'==================================================================================================
Function コメント削除()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  If ActiveSheet.ProtectContents = True Then
    Call Library.showNotice(413, , True)
  End If
  For Each slctCells In Selection
    If TypeName(slctCells.Comment) = "Comment" Then
      slctCells.ClearComments
    End If
    DoEvents
  Next
  
  '処理終了--------------------------------------
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
  Call Library.errorHandle
End Function

'==================================================================================================
Function 行例を入れ替えて貼付け()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
'  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

  
  '処理終了--------------------------------------
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
  Call Library.errorHandle
End Function


'==================================================================================================
Function ゼロ埋め()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.ゼロ埋め"
  
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
  
  On Error Resume Next
  Selection.SpecialCells(xlCellTypeBlanks).Value = 0
  On Error GoTo catchError

  
  '処理終了--------------------------------------
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
  Call Library.errorHandle
End Function

'==================================================================================================
Function 指定フォントに設定()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.指定フォントに設定"
  
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
  
  On Error Resume Next
  Selection.SpecialCells(xlCellTypeBlanks).Value = 0
  On Error GoTo catchError

  
  '処理終了--------------------------------------
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
  Call Library.errorHandle
End Function


'==================================================================================================
Function 改行削除()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.Trim01"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctCells In Selection
    resVal = slctCells.Text
    
    If resVal <> "" Then
      resVal = Replace(resVal, vbCrLf, "")
      resVal = Replace(resVal, vbCr, "")
      resVal = Replace(resVal, vbLf, "")
      slctCells.Value = resVal
      DoEvents
    End If
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 行挿入()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.行挿入"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  'Set slctCells = Selection
  
  Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 列挿入()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.列挿入"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  'Set slctCells = Selection
  
  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Application.CutCopyMode = False


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 定数削除()
  Dim slctCells As Range
  Dim resVal As String
  Const funcName As String = "Ctl_Cells.定数削除"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  On Error Resume Next
  If Selection.count = 1 Then
    Call Library.showNotice(600)
  ElseIf Selection.count > 1 Then
    Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents
  End If
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.セル幅調整"
  
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
  If Selection.count > 1 Then
    Columns(Library.getColumnName(Selection(1).Column) & ":" & Library.getColumnName(Selection(Selection.count).Column)).EntireColumn.AutoFit
    
    For Each slctCells In Selection
      colName = Library.getColumnName(slctCells.Column)
      If IsNumeric(slctCells.Value) Then
        If CInt(slctCells.Value) > 1 Then
          Columns(colName & ":" & colName).ColumnWidth = slctCells.Value
        End If
      End If
      
      If Columns(colName & ":" & colName).ColumnWidth > 60 Then
        Columns(colName & ":" & colName).ColumnWidth = 60
      End If
    Next

  Else
    Cells.EntireColumn.AutoFit
    For colLine = 1 To Columns.count
      colName = Library.getColumnName(colLine)
      If IsNumeric(Cells(1, colLine)) Then
        If CInt(Cells(1, colLine)) > 1 Then
          Columns(colName & ":" & colName).ColumnWidth = Cells(1, colLine).Value
        End If
      End If
      
      If Cells(1, colLine).ColumnWidth > 60 Then
        colName = Library.getColumnName(colLine)
        Columns(colName & ":" & colName).ColumnWidth = 60
      End If
    Next
  End If
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript(True)
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
  Call Library.errorHandle
End Function

'==================================================================================================
Function セル高さ調整()
  Const funcName As String = "Ctl_Cells.セル高さ調整"
  
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
  
  Cells.EntireRow.AutoFit
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript(True)
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
