Attribute VB_Name = "Ctl_shap"
Option Explicit


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @Link https://infoment.hatenablog.com/entry/2021/08/17/000649
'**************************************************************************************************
Function TextToFitShape(targetShape As Excel.Shape, Optional chkFlg As Boolean = True) As Long
  ' テキストの有無確認。無い場合は、Functionを終了する。
  If targetShape.TextFrame2.TextRange.Characters.Text = vbNullString Then
      Exit Function
  End If

  ' オートシェイプのサイズ取得。
  Dim h(1) As Double: h(0) = targetShape.Height
  Dim w(1) As Double: w(0) = targetShape.Width
  Dim l As Double: l = targetShape.Left
  Dim T As Double: T = targetShape.Top
  
  ' オートシェイプを一旦、文字サイズに合わせてサイズ変更。
  targetShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
  
  ' 変更後のサイズ取得。
  h(1) = targetShape.Height
  w(1) = targetShape.Width
  
  ' オートシェイプの縦と横、各々の縮小（もしくは拡大）率のうち、
  ' 小さい方を取得（大きい方だと、オートシェイプから食み出る）。
  Dim ρ As Double
  ρ = WorksheetFunction.Min(h(0) / h(1), w(0) / w(1))
  
  ' もとのフォントサイズにρを掛け、目安のフォントサイズを得る。
  Dim FontSize As Long
  FontSize = targetShape.TextFrame2.TextRange.Font.Size * ρ
      
  Dim i As Long
  Do
    ' フォントサイズ仮決め。
    targetShape.TextFrame2.TextRange.Font.Size = FontSize
    
    ' 改めて、オートシェイプを文字サイズに合わせてサイズ変更。
    targetShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    
    ' 変更後のサイズを得る。
    h(1) = targetShape.Height
    w(1) = targetShape.Width
    
    ' 縦と横どちらか一方でも元のサイズを越えたら、そこで終了。
    If (h(1) > h(0) Or w(1) > w(0)) And chkFlg = True Then
      Exit Do
    
    ElseIf (w(1) > w(0)) And chkFlg = False Then
      Exit Do
    
    ' そうでなければ、まだピッタリではない。フォントサイズを１増加。
    Else
        FontSize = FontSize + 1
    End If
    
    ' 無限ループ防止。
    i = i + 1: If i >= 100 Then Exit Do
  Loop
  
  ' サイズを越えてから抜けたので、１引いて丁度のサイズにする。
  FontSize = FontSize - 1
  
  ' オートサイズ解除。
  targetShape.TextFrame2.AutoSize = msoAutoSizeNone
  
  ' オートシェイプを最初の大きさに戻す。
  targetShape.Height = h(0)
  targetShape.Width = w(0)
  
  targetShape.Left = l
  targetShape.Top = T
  
  ' フォントサイズを最終値に変更。
  targetShape.TextFrame2.TextRange.Font.Size = FontSize
  
  ' 戻り値としてフォントサイズを返す。
  TextToFitShape = FontSize
End Function



'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************



'==================================================================================================
Function QRコード生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range, targetCells As Range
  
  Dim chartAPIURL As String
  Dim QRCodeImgName As String
  Dim colSize As Long, colHeight As Long, colWidth As Long
  
  Const funcName As String = "Ctl_Shap.QRコード生成"
  Const chartAPI = "https://chart.googleapis.com/chart?cht=qr&chld=l|1&"
  
  
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

  With Frm_mkQRCode
    .Show
  End With
  
  
  For Each slctCells In Selection
    QRCodeImgName = "QRCode_" & slctCells.Address(False, False)
    
    '既存を削除
    If Library.chkShapeName(QRCodeImgName) Then
      ActiveSheet.Shapes.Range(Array(QRCodeImgName)).Select
      Selection.delete
    End If
    
    colHeight = FrmVal("codeSize") * 0.75 + 4
    colWidth = FrmVal("codeSize") * 0.118 + 4
    Set targetCells = Range(FrmVal("CellAddress") & slctCells.Row)
    
    With targetCells
      .Select
      If FrmVal("onReSize") = True Then
        If .rowHeight < colHeight Then .rowHeight = colHeight
        If .ColumnWidth < colWidth Then .ColumnWidth = colWidth
      End If
    End With
    
    chartAPIURL = chartAPI & "chs=" & FrmVal("codeSize") & "x" & FrmVal("codeSize")
    chartAPIURL = chartAPIURL & "&chl=" & Library.convURLEncode(slctCells.Text)
    
    Call Library.showDebugForm("chartAPIURL", chartAPIURL, "debug")
    
    With ActiveSheet.Pictures.Insert(chartAPIURL)
      If FrmVal("onReSize") = True Then
        .ShapeRange.Top = targetCells.Top + (targetCells.Height - .ShapeRange.Height) / 2
        .ShapeRange.Left = targetCells.Left + (targetCells.Width - .ShapeRange.Width) / 2
      Else
        .ShapeRange.Top = targetCells.Top
        .ShapeRange.Left = targetCells.Left
      End If
      
      .Placement = xlMove
      
      'QRコードの名前設定
      .ShapeRange.Name = QRCodeImgName
      .Name = QRCodeImgName
    
    End With
    DoEvents
    Set targetCells = Nothing
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
Function パスワード生成()
  Const funcName As String = "Ctl_Shap.パスワード生成"


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
  '----------------------------------------------

  Frm_mkPasswd.Show vbModeless


  '処理終了--------------------------------------
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
Function サイズ合わせ_間隔設定()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseHeight As Long, baseWidth As Long, baseTop As Long
  Dim baseShape As Shape, shp As Shape, beforeShape As Shape
  Dim shCnt As Long

  Const funcName As String = "Ctl_Shap.サイズ合わせ_間隔設定"
  Const interval As Long = 10
  
  
  
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

  If VarType(Selection) <> vbObject Then
    GoTo LB_ExitFunction
  End If

  Set baseShape = Selection.ShapeRange.Item(1)
  baseHeight = CLng(baseShape.Height)
  baseWidth = CLng(baseShape.Width)
  baseTop = CLng(baseShape.Top)

  baseWidth = 236

  For shCnt = 1 To Selection.ShapeRange.count
    Set shp = Selection.ShapeRange.Item(shCnt)
  
    shp.Height = baseHeight
    shp.Width = baseWidth
    shp.Top = baseTop
    
    If shCnt > 1 Then
      Set beforeShape = Selection.ShapeRange.Item(shCnt - 1)
      shp.Left = beforeShape.Left + beforeShape.Width + interval
    End If
  Next

  '処理終了--------------------------------------
LB_ExitFunction:
  If runFlg = False Then
    
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function サイズ合わせ_幅()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseHeight As Long, baseWidth As Long, baseTop As Long
  Dim baseShape As Shape, shp As Shape, beforeShape As Shape
  Dim shCnt As Long

  Const funcName As String = "Ctl_Shap.サイズ合わせ_幅"
  Const interval As Long = 10
  
  
  
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

  If VarType(Selection) <> vbObject Then
    GoTo LB_ExitFunction
  End If

  Set baseShape = Selection.ShapeRange.Item(1)
  baseHeight = CLng(baseShape.Height)
  baseWidth = CLng(baseShape.Width)
  
  For shCnt = 1 To Selection.ShapeRange.count
    Set shp = Selection.ShapeRange.Item(shCnt)
  
    'shp.Height = baseHeight
    shp.Width = baseWidth
  Next

  '処理終了--------------------------------------
LB_ExitFunction:
  If runFlg = False Then
    
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function サイズ合わせ_高さ()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseHeight As Long, baseWidth As Long, baseTop As Long
  Dim baseShape As Shape, shp As Shape, beforeShape As Shape
  Dim shCnt As Long

  Const funcName As String = "Ctl_Shap.サイズ合わせ_高さ"
  Const interval As Long = 10
  
  
  
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

  If VarType(Selection) <> vbObject Then
    GoTo LB_ExitFunction
  End If

  Set baseShape = Selection.ShapeRange.Item(1)
  baseHeight = CLng(baseShape.Height)
  baseWidth = CLng(baseShape.Width)
  
  For shCnt = 1 To Selection.ShapeRange.count
    Set shp = Selection.ShapeRange.Item(shCnt)
  
    shp.Height = baseHeight
    'shp.Width = baseWidth
  Next

  '処理終了--------------------------------------
LB_ExitFunction:
  If runFlg = False Then
    
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, Err.Description, "Error")
  Call Library.errorHandle
End Function

