Attribute VB_Name = "Ctl_RbnMaint"
Option Explicit

'**************************************************************************************************
' * Copy
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function シート追加()
  Call init.setting
  ThisWorkbook.Worksheets.add.Name = "inputData"
  ThisWorkbook.Save
End Function


'==================================================================================================
Function シート削除()
  Call init.setting
  Application.DisplayAlerts = False
  
  ThisWorkbook.Sheets("HighLight").delete
  
  Application.DisplayAlerts = True
  ThisWorkbook.Save
End Function


'==================================================================================================
Function その他()
  Dim line As Long, endLine As Long
  
  Call init.setting(True)
  BK_ThisBook.Sheets("Help").Activate
'  Cells.Select
'  Selection.ColumnWidth = 5
  
'  LadexSh_HiLight.Range("N5").Select
'  LadexSh_HiLight.Range("N5").ClearComments
'  LadexSh_HiLight.Range("N5").AddComment
'  LadexSh_HiLight.Range("N5").Comment.Visible = False
'  LadexSh_HiLight.Range("N5").Comment.Text Text:="Sampleコメント" & Chr(10) & "Sampleコメント" & Chr(10) & "Sampleコメント "
'
'  LadexSh_HiLight.Range("N5").Comment.Visible = True
  
  
  LadexSh_Help.Cells.ColumnWidth = 3
  LadexSh_Help.Cells.rowHeight = 15
  
  endLine = LadexSh_Help.Cells(Rows.count, 1).End(xlUp).Row
  For line = 1 To endLine
    If LadexSh_Help.Range("A" & line) <> "" Then
      LadexSh_Help.Cells.rowHeight = 20
    End If
  Next
  
  
  ThisWorkbook.Save
End Function

'==================================================================================================
Function OptionSheetImport()
  Dim line As Long, endLine As Long
  Dim objShp
  
  Call init.setting
  Call Ctl_ProgressBar.showStart
  
  Set targetBook = Workbooks("メンテナンス用.xlsm")

  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 1, 10, "設定")
  targetBook.Sheets("設定").Columns("A:Z").Copy ThisWorkbook.Worksheets("設定").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 2, 10, "Notice")
  targetBook.Sheets("Notice").Columns("A:Z").Copy ThisWorkbook.Worksheets("Notice").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 3, 10, "Style")
  targetBook.Sheets("Style").Columns("A:Z").Copy ThisWorkbook.Worksheets("Style").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 4, 10, "testData")
  targetBook.Sheets("testData").Columns("A:AZ").Copy ThisWorkbook.Worksheets("testData").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 5, 10, "Favorite")
  targetBook.Sheets("Favorite").Columns("A:Z").Copy ThisWorkbook.Worksheets("Favorite").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 6, 10, "Function")
  targetBook.Sheets("Function").Columns("A:Z").Copy ThisWorkbook.Worksheets("Function").Range("A1")
  
  
  Application.DisplayAlerts = False
  
  'ハイライト、コメントプレビュー用--------------
'  ThisWorkbook.Sheets("HighLight").delete
'  ThisWorkbook.Worksheets.add.Name = "HighLight"
'  ThisWorkbook.Sheets("HighLight").Cells.ColumnWidth = 3.86
'  ThisWorkbook.Sheets("HighLight").Cells.RowHeight = 15
'  targetBook.Sheets("HighLight").Columns("A:Z").Copy ThisWorkbook.Worksheets("HighLight").Range("A1")
  
  
  'ヘルプシート編集------------------------------
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 7, 10, "Help")

  ThisWorkbook.Sheets("Help").Cells.ColumnWidth = 3
  ThisWorkbook.Sheets("Help").Cells.rowHeight = 15
  
  ThisWorkbook.Worksheets("Help").Cells.delete Shift:=xlUp
  For Each objShp In ThisWorkbook.Worksheets("Help").Shapes
    objShp.delete
  Next
  targetBook.Sheets("Help").Columns("A:AZ").Copy ThisWorkbook.Worksheets("Help").Range("A1")
  
  'スタンプシート編集----------------------------
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 7, 10, "Stamp")

  ThisWorkbook.Sheets("Stamp").delete
  ThisWorkbook.Worksheets.add.Name = "Stamp"
  targetBook.Sheets("Stamp").Columns("A:AP").Copy ThisWorkbook.Worksheets("Stamp").Range("A1")
    
 
  Application.DisplayAlerts = True
  ThisWorkbook.Save
  Set targetBook = Nothing
  
  Call Ctl_ProgressBar.showEnd
End Function


'==================================================================================================
Function OptionSheetExport()

  Call init.setting
  Call Ctl_ProgressBar.showStart

  Set targetBook = Workbooks("メンテナンス用.xlsm")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 1, 10, "設定")
  ThisWorkbook.Sheets("設定").Columns("A:AA").Copy targetBook.Worksheets("設定").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 2, 10, "Notice")
  ThisWorkbook.Sheets("Notice").Columns("A:B").Copy targetBook.Worksheets("Notice").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 3, 10, "Style")
  ThisWorkbook.Sheets("Style").Columns("A:J").Copy targetBook.Worksheets("Style").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 4, 10, "testData")
  ThisWorkbook.Sheets("testData").Columns("A:P").Copy targetBook.Worksheets("testData").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 5, 10, "Favorite")
  ThisWorkbook.Worksheets("Favorite").Columns("A:C").Copy targetBook.Sheets("Favorite").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 6, 10, "Function")
  ThisWorkbook.Worksheets("Function").Columns("A:Z").Copy targetBook.Sheets("Function").Range("A1")
  
  Call Ctl_ProgressBar.showBar("メンテナンス", 1, 2, 7, 10, "SheetList")
  ThisWorkbook.Worksheets("SheetList").Columns("A:Z").Copy targetBook.Sheets("SheetList").Range("A1")
  
'  Call Library.showDebugForm(ThisWorkbook.Worksheets("Ribbon").Range("A2"))
  
  targetBook.Save
  ThisWorkbook.Save
  Set targetBook = Nothing
  Call Ctl_ProgressBar.showEnd
End Function


