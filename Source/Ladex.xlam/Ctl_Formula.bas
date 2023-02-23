Attribute VB_Name = "Ctl_Formula"
Option Explicit

'**************************************************************************************************
' * 数式内のセル参照
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @Link https://excel-ubara.com/excelvba5/EXCELVBA258.html
'**************************************************************************************************
'==================================================================================================
Function 数式確認()
  Dim confirmFormulaName As String
  Dim count As Long
  Dim formulaVals As Variant
  Dim objShp, aryRange
  Const funcName As String = "Ctl_Formula.数式確認"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  '既存のオブジェクト削除
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  If ActiveCell.HasFormula = False Or BKcf_rbPressed = False Then
    Call Library.endScript
    Exit Function
  End If
  aryRange = Ctl_Formula.getFormulaRange(ActiveCell)
  
  count = 1
  For Each formulaVals In aryRange
    confirmFormulaName = "confirmFormulaName_" & count
    Call Ctl_Formula.範囲選択(formulaVals, confirmFormulaName)
    count = count + 1
  Next
  ActiveCell.Select
  
  '処理終了--------------------------------------
  Call Library.endScript
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 範囲選択(formulaVals As Variant, confirmFormulaName As String)
  Const funcName As String = "Ctl_Formula.範囲選択"
  
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
  
  If formulaVals.Worksheet.Name <> ActiveSheet.Name Then
    Exit Function
  End If

  With ActiveSheet.Range(formulaVals.Address(external:=False))
    ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height).Select
  End With
  
  Selection.Name = confirmFormulaName
  Selection.ShapeRange.Fill.ForeColor.RGB = RGB(205, 205, 255)
  Selection.ShapeRange.Fill.Transparency = 0.5
  Selection.OnAction = "Ctl_Formula.GetCurPosition"
  Selection.Text = formulaVals.Address(RowAbsolute:=False, ColumnAbsolute:=False, external:=False)
  
  With Selection.ShapeRange.TextFrame2
    .TextRange.Font.NameComplexScript = "メイリオ"
    .TextRange.Font.NameFarEast = "メイリオ"
    .TextRange.Font.Name = "メイリオ"
    .TextRange.Font.Size = 9
    .MarginLeft = 3
    .MarginRight = 0
    .MarginTop = 0
    .MarginBottom = 0
    .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
  End With
  
  Selection.ShapeRange.line.Visible = msoTrue
  Selection.ShapeRange.line.ForeColor.RGB = RGB(255, 0, 0)
  Selection.ShapeRange.line.Weight = 2
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function GetCurPosition()
  Dim p As POINTAPI 'API用変数
  Dim Rng  As Range
  Dim objShp
  Const funcName As String = "Ctl_Formula.GetCurPosition"
  
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
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  Call Library.waitTime(50)
  
  'カーソル位置取得
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.y).Select
  End If
  Call Ctl_Formula.数式確認
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function getFormulaRange(ByVal argRange As Range) As Range()
  Dim sFormula As String, sSplit() As String, sTemp As String
  Dim aryRange() As Range, tRange As Range
  Dim ix As Long, i As Long
  Dim flgS As Boolean, flgD As Boolean
  Const funcName As String = "Ctl_Formula.getFormulaRange"
  
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
  '=以降の計算式
  sFormula = Mid(argRange.FormulaLocal, 2)
  '計算式の中の改行や余分な空白を除去
  sFormula = Replace(sFormula, vbCrLf, "")
  sFormula = Replace(sFormula, vbLf, "")
  sFormula = Trim(sFormula)

  flgS = False
  flgD = False
  For i = 1 To Len(sFormula)
    'シングル・ダブルのTrue,Falseを反転
    Select Case Mid(sFormula, i, 1)
      Case "'"
        flgS = Not flgS
      Case """"
        'シングルの中ならシート名
        If Not flgS Then
          flgD = Not flgD
        End If
    End Select
    Select Case Mid(sFormula, i, 1)
      '各種演算子の判定
      Case "+", "-", "*", "/", "^", ">", "<", "=", "(", ")", "&", ",", " "
        Select Case True
          Case flgS
            'シングルの中ならシート名
            sTemp = sTemp & Mid(sFormula, i, 1)
          Case flgD
            'ダブルの中なら無視
          Case Else
            '各種演算子をvbLfに置換
            sTemp = sTemp & vbLf
        End Select
      Case Else
        'ダブルの中なら無視、ただしシングルの中はシート名
        If Not flgD Or flgS Then
          sTemp = sTemp & Mid(sFormula, i, 1)
        End If
    End Select
  Next

  On Error Resume Next
  'vbLfで区切って配列化
  sSplit = Split(sTemp, vbLf)
  ix = 0
  For i = 0 To UBound(sSplit)
    If sSplit(i) <> "" Then
      Err.Clear
      'Application.Evaluateメソッドを使ってRangeに変換
      If InStr(sSplit(i), "!") > 0 Then
        Set tRange = Evaluate(Trim(sSplit(i)))
      Else
        'シート名を含まない場合は、元セルのシート名を付加
        Set tRange = Evaluate("'" & argRange.Parent.Name & "'!" & Trim(sSplit(i)))
      End If
      'Rangeオブジェクト化が成功すれば配列へ入れる
      If Err.Number = 0 Then
        ReDim Preserve aryRange(ix)
        Set aryRange(ix) = tRange
        ix = ix + 1
      End If
    End If
  Next
  On Error GoTo 0
  getFormulaRange = aryRange
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * 数式編集
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function エラー防止_空白()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.エラー防止_空白"
  
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
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IFERROR(" & formulaVal & ","""")"
      
      slctCells.Formula = "=" & formulaVal
    End If
  Next
  
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function エラー防止_ゼロ()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.エラー防止_ゼロ"
  
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
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IFERROR(" & formulaVal & ",0)"
      
      slctCells.Formula = "=" & formulaVal
    End If
  Next
  
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ゼロ非表示()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.ゼロ非表示"
  
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
  
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IF(" & formulaVal & "=0,""""," & formulaVal & ")"
      
      slctCells.Formula = "=" & formulaVal
    End If
  Next
  
  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 行番号追加()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Formula.行番号追加"

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
  
  line = Selection.Row - 1
  For Each slctCells In Selection
    slctCells.FormulaR1C1 = "=ROW()-" & line
    DoEvents
  Next

  '処理終了--------------------------------------
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 数式挿入(formulaType As String)
  Dim confirmFormulaName As String
  Dim count As Long
  Dim formulaVals As Variant
  Dim objShp, aryRange
  Const funcName As String = "Ctl_Formula.数式挿入"
  
  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.showDebugForm("formulaType", formulaType, "debug")
  '----------------------------------------------
  
  Select Case formulaType
    Case "SheetName"
      ActiveCell.FormulaR1C1 = "=MID(CELL(""filename"",RC),FIND(""]"",CELL(""filename"",RC))+1,31)"


    Case Else
  End Select

  
  '処理終了--------------------------------------
  Call Library.endScript
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
