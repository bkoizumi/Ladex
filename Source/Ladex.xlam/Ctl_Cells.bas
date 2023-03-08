Attribute VB_Name = "Ctl_Cells"
Option Explicit

'**************************************************************************************************
' * セル操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function セル高さ固定設定(setRowHeight As Integer)
  Dim line As Long, startLine As Long, endLine As Long
  Dim SelectionCell As Range
  
  Const funcName As String = "Ctl_Cells.セル高さ固定設定"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")
  Call Library.showDebugForm("setRowHeight", setRowHeight, "debug")

  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  startLine = Selection(1).Row
  endLine = Selection(Selection.count)
  If endLine = 0 Then
    endLine = startLine
  End If
  Rows(startLine & ":" & endLine).rowHeight = setRowHeight
  
  
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function





'==================================================================================================
Function 前後のスペース削除()
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.前後のスペース削除"

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

  'エラー発生時------------------------------------------------------------------------------------
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 中黒点付与()
  Dim line As Long, endLine As Long
  Dim Reg As Object
  Dim slctCells
  
  Const funcName As String = "Ctl_Cells.中黒点付与"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 連番設定()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim i As Long
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.連番設定"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 連番追加()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  Const funcName As String = "Ctl_Cells.連番追加"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 英数字全⇒半角変換()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.英数字全⇒半角変換"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 英数字半⇒全角変換()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim i As Long
  Dim arrCells
  Const funcName As String = "Ctl_Cells.英数字半⇒全角変換"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 取り消し線設定()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.取り消し線設定"
    
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function コメント挿入()
  Dim commentVal As String, commentBgColor As Long, CommentFontColor As Long
  Dim CommentFont As String, CommentFontSize As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント挿入"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function コメント削除()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.コメント削除"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 行例を入れ替えて貼付け()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.行例を入れ替えて貼付け"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ゼロ埋め()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.ゼロ埋め"
  
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 改行削除()
  Dim slctCells As Range
  Dim resVal As String
  Dim i As Long
  
  Const funcName As String = "Ctl_Cells.改行削除"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 行挿入()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.行挿入"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 列挿入()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.列挿入"

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

  'エラー発生時------------------------------------------------------------------------------------
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function セル幅調整()
  Dim colLine As Long, endColLine As Long
  Dim slctStartColLine As Long, slctEndColLine As Long, useEndColLine As Long
  Dim colName As String
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Cells.セル幅調整"
  Const maxColumnWidth As Integer = 60
  
  
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
  
  useEndColLine = Selection.SpecialCells(xlLastCell).Column
  
  '選択範囲がある場合----------------------------
  If Selection.CountLarge > 1 Then
    Call Library.showDebugForm("選択範囲あり", , "debug")
    
    Selection.EntireColumn.AutoFit
  
    slctStartColLine = Selection.Column
    slctEndColLine = Selection.Column + Selection.Columns.count
  
  '選択範囲がない場合----------------------------
  Else
    Call Library.showDebugForm("選択範囲なし", , "debug")
    Cells.EntireColumn.AutoFit
  
    slctStartColLine = 1
    slctEndColLine = useEndColLine
  
  End If
  
  '最大幅確認------------------------------------
  For colLine = slctStartColLine To slctEndColLine
    colName = Library.getColumnName(colLine)
    Call Library.showDebugForm("最大幅確認", colName, "debug")
    
    
    If Cells(1, colLine).ColumnWidth > maxColumnWidth Then
      Columns(colName & ":" & colName).ColumnWidth = maxColumnWidth
    Else
      Columns(colName & ":" & colName).ColumnWidth = WorksheetFunction.RoundUp(Columns(colName & ":" & colName).ColumnWidth, 0) + 1
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function セル高さ調整()
  Dim line As Long, startLine As Long, endLine As Long
  Dim SelectionCell As Range
  
  Const funcName As String = "Ctl_Cells.セル高さ調整"
  
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
  
  Cells.EntireRow.AutoFit
  
  If Selection.Rows.count <= 1 Then
    startLine = 1
    endLine = Range("A1").SpecialCells(xlLastCell).Row
  Else
    startLine = SelectionCell.Row
    endLine = Range("A1").SpecialCells(xlLastCell).Row
  End If
  Selection.EntireRow.AutoFit
  
  
  For line = startLine To endLine
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "高さ設定")
    
    If Rows(line & ":" & line).Hidden = False Then
      If Rows(line & ":" & line).Height < Int(dicVal("rowHeight")) Then
        Rows(line & ":" & line).rowHeight = dicVal("rowHeight")
      
      ElseIf Rows(line & ":" & line).Height > 200 Then
        Rows(line & ":" & line).rowHeight = 200
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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 大⇒小変変換()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.大⇒小変変換"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 小⇒大変変換()
  Dim line As Long, endLine As Long
  Dim slctCellsCnt As Long
  Dim arrCells
  
  Const funcName As String = "Ctl_Cells.小⇒大変変換"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 丸数字⇒数値()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Const funcName As String = "Ctl_Cells.丸数字⇒数値"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 数値⇒丸数字()
  Dim line As Long, endLine As Long
  Dim slctCells As Range

  Const funcName As String = "Ctl_Cells.数値⇒丸数字"

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
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function URLエンコード()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.URLエンコード"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function URLデコード()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.URLデコード"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function Unicodeエスケープ()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.Unicodeエスケープ"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function Unicodeアンエスケープ()
  Dim slctCells As Range
  Const funcName As String = "Ctl_Cells.Unicodeアンエスケープ"

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

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
