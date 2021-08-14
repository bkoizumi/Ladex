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
  
'  On Error GoTo catchError
  Call Library.startScript
  
  '既存のファイル削除
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  If ActiveCell.HasFormula = False Or BKcf_rbPressed = False Then
    Call Library.endScript
    Exit Function
  End If
  
  Call init.setting
  aryRange = getFormulaRange(ActiveCell)
  
  count = 1
  For Each formulaVals In aryRange
    confirmFormulaName = "confirmFormulaName_" & count
  
    Call 範囲選択(formulaVals, confirmFormulaName)
    count = count + 1
  Next
  
  ActiveCell.Select


  Call Library.endScript
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
  'Call Library.showNotice(400, "", True)
End Function


'==================================================================================================
Function 範囲選択(formulaVals As Variant, confirmFormulaName As String)

  If formulaVals.Worksheet.Name <> ActiveSheet.Name Then
    Exit Function
  End If

  With ActiveSheet.Range(formulaVals.Address(external:=False))
    ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, width:=.width, height:=.height).Select
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
  
  
End Function



'==================================================================================================
Sub GetCurPosition()

  Dim p        As POINTAPI 'API用変数
  Dim Rng  As Range
  Dim objShp
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  Call Library.waitTime(50)
  Call Library.startScript
  
  'カーソル位置取得
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.Y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.Y).Select
  End If
  Call Library.endScript
  Call Ctl_Formula.数式確認


End Sub

'==================================================================================================
Function getFormulaRange(ByVal argRange As Range) As Range()
    Dim sFormula As String
    Dim aryRange() As Range
    Dim tRange As Range
    Dim ix As Long
    Dim i As Long
    Dim flgS As Boolean 'シングルクオートが奇数の時True
    Dim flgD As Boolean 'ダブルクオートが奇数の時True
    Dim sSplit() As String
    Dim sTemp As String
  
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
End Function



'**************************************************************************************************
' * 数式編集
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function formula01()
  Dim formulaVal As String
  
  'On Error GoTo catchError
  
  If ActiveCell.HasFormula = False Then
    Exit Function
  End If
  Call init.setting
  
  formulaVal = ActiveCell.Formula
  formulaVal = Replace(formulaVal, "=", "")
  formulaVal = Replace(formulaVal, vbCrLf, "")
  formulaVal = Replace(formulaVal, vbLf, "")
  formulaVal = Trim(formulaVal)
  
  formulaVal = "IFERROR(" & formulaVal & ","""")"
  
  ActiveCell.Formula = "=" & formulaVal
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function

'==================================================================================================
Function formula02()
  Dim formulaVal As String
  
  'On Error GoTo catchError

  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function


'==================================================================================================
Function formula03()
  Dim formulaVal As String
  
  'On Error GoTo catchError

  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function







