Attribute VB_Name = "StyleSetting"
Function StyleSetting_Sheet()

  Dim line As Integer
 
  Columns("A:A").ColumnWidth = 1
  Columns("B:B").ColumnWidth = 5
  Columns("C:C").ColumnWidth = 15
  Columns("D:D").ColumnWidth = 25
  Columns("E:E").ColumnWidth = 25
  Columns("F:F").ColumnWidth = 10
  Columns("G:G").ColumnWidth = 13.29
  Columns("H:H").ColumnWidth = 11.13
  Columns("J:J").ColumnWidth = 2.25
  Columns("K:K").ColumnWidth = 2.25
  Columns("L:L").ColumnWidth = 2.25
  Columns("M:M").ColumnWidth = 2.25
  Columns("N:N").ColumnWidth = 2.25
  Columns("O:O").ColumnWidth = 11.38
  Columns("P:U").ColumnWidth = 12.75

  For line = 9 To 108
    Range("C" & line).Value = StyleSetting_han2zen(Range("C" & line).Value)
    Range("Q" & line).Value = StyleSetting_han2zen(Range("Q" & line).Value)
    
    Range("O" & line).Value = Replace(Range("O" & line).Value, "NOT NULL", "1")
    
    If Range("G" & line).Value = "numeric" Then
      If InStr(Range("H" & line).Value, ",") = 0 Then
        Range("H" & line).Select
        Selection.Style = "注意"
      End If
    End If
    
    
  Next line

End Function


Function StyleSetting_Cell()


  Selection.RowHeight = Selection.RowHeight + 20





End Function



Sub 印刷範囲設定()

  On Error GoTo ErrHand
  
  Dim endLine As Integer
  Dim PageCnt As Integer
  Dim OnePageRow As Integer
  Dim RowCnt As Integer
  Dim ThisActiveSheetName As String
  Dim WindowZoomLevel As Integer
  
  WindowZoomLevel = ActiveWindow.Zoom
  
  ThisActiveSheetName = activeSheet.Name
  
  endLine = activeSheet.Cells(Rows.count, 2).End(xlUp).Row
  OnePageRow = 30
  PageCnt = 1
  
  ' ======================= 処理開始 ======================
  '改ページプレビュー
  ActiveWindow.View = xlPageBreakPreview
  
  'すべての改ページを解除
  activeSheet.ResetAllPageBreaks
  
  '印刷範囲をクリアする
  activeSheet.PageSetup.PrintArea = ""
  
  '印刷範囲の詳細設定
  With activeSheet.PageSetup
    .RightFooter = "&""Arial,標準""&8Sharp Business Solutions Corporation"
    .CenterFooter = "&P / &N"
'    .PrintTitleRows = "$2:$8"                 '行タイトル
    .PrintArea = "$B$2:$U$" & endLine
    .BlackAndWhite = False                    '白黒印刷 True:する  False:しない
    .Zoom = False                             '拡大・縮小率を指定しない
    .FitToPagesTall = False                   '縦方向は指定しない
    .FitToPagesWide = 1                       '横方向1ページで印刷
    
    .TopMargin = Application.CentimetersToPoints(1.5)       '上余白
    .BottomMargin = Application.CentimetersToPoints(1.5)    '下余白
    .LeftMargin = Application.CentimetersToPoints(1)        '左余白
    .RightMargin = Application.CentimetersToPoints(1)       '右余白
    .HeaderMargin = Application.CentimetersToPoints(0.8)    'ヘッダー余白
    .FooterMargin = Application.CentimetersToPoints(0.7)    'フッター余白
  End With

  '標準画面に戻す
'  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = 80
  Range("A1").Select
Exit Sub

ErrHand:
  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = WindowZoomLevel
End Sub


Function StyleSetting_han2zen(Text As String)


  Dim c As Range
  Dim i As Integer
  Dim rData As Variant, ansData As Variant

  ansData = ""
  
  For i = 1 To Len(Text)
    rData = StrConv(Text, vbWide)
    If Mid(rData, i, 1) Like "[Ａ-ｚ]" Or Mid(rData, i, 1) Like "[０-９]" Or Mid(rData, i, 1) Like "－" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
     ansData = ansData & Mid(rData, i, 1)
    End If
    
    
    ansData = Replace(ansData, "（", "(")
    ansData = Replace(ansData, "）", ")")
    ansData = Replace(ansData, ":", "：")
    ansData = Replace(ansData, "::", "：")
    ansData = Replace(ansData, "()", "")
    'ansData = Replace(ansData, "　", "、")

  Next i
 
  StyleSetting_han2zen = ansData

End Function
