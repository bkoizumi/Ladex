Attribute VB_Name = "Ctl_Help"

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

  Dim actrow As Integer
  On Error GoTo catchError


  ' 選択セルが変更されたとき
  If ActiveCell.Column = 1 And ActiveCell.Value <> "" Then
    ' A列で値が"タイトル"で選択範囲が3の場合そのセルを左上に持ってくる
    With ActiveWindow
      .ScrollRow = Target.Row
      .ScrollColumn = Target.Column
    End With
  End If
  Exit Sub

'---------------------------------------------------------------------------------------
'エラー発生時の処理
'---------------------------------------------------------------------------------------
catchError:

End Sub



'**************************************************************************************************
' * 目次生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub 目次生成()

  Dim line As Long, endLine As Long, mline As Long
  Dim columnName1 As String, columnName2 As String

  On Error GoTo catchError:

  mline = 2

'  ActiveWorkbook.Worksheets("Help").Select
  endLine = Cells(Rows.count, 1).End(xlUp).Row

  Range("B2:AA28").ClearContents
  Range("AB1") = Format(Date, "yyyy/mm/dd")

  For line = 30 To endLine
    If Range("A" & line) <> "" Then
      If Range("A" & line) = "5．運用手順" Then
        mline = 2
      End If

      If Range("A" & line) Like "5*" Then
        columnName1 = "P"
        columnName2 = "Z"
      Else
        columnName1 = "B"
        columnName2 = "L"
      End If

      With Range(columnName1 & mline)
        .Value = Range("A" & line)
        .Select
        .Hyperlinks.add anchor:=Selection, Address:="", SubAddress:="#" & "A" & line
        .Font.ColorIndex = 1
        .Font.Underline = xlUnderlineStyleNone
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .Font.Name = "メイリオ"
        .Font.Size = 9
        .Font.Bold = True
      End With
      Range(columnName1 & mline & ":" & columnName2 & mline).Select
      'Selection.Merge
      With Selection
          .Merge
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .IndentLevel = 0
          .ShrinkToFit = True
          .ReadingOrder = xlContext
          .MergeCells = True

          If Range("A" & line) Like "*-*" Then
            .InsertIndent 2
            .Font.Bold = False
          End If
      End With
      mline = mline + 1
    End If

  Next
  Application.GoTo Reference:=Range("A1"), Scroll:=True


  Exit Sub
'---------------------------------------------------------------------------------------
'エラー発生時の処理
'---------------------------------------------------------------------------------------
catchError:

    Call Library.endScript

End Sub
