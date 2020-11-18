Attribute VB_Name = "Copy"
'**************************************************************************************************
' * Copy
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function データコピー()
  
  Set targetBook = Workbooks("元BK_Library.xlsx")
  
'  ThisWorkbook.Sheets("Ribbon").Columns("A:G").Value = targetBook.Sheets("Ribbon").Columns("A:G").Value
'  ThisWorkbook.Sheets("Notice").Columns("A:B").Value = targetBook.Sheets("Notice").Columns("A:B").Value
'  ThisWorkbook.Worksheets("Style").Columns("A:J").Value = targetBook.Sheets("Style").Columns("A:J").Value
'
  targetBook.Sheets("Ribbon").Columns("A:G").Copy ThisWorkbook.Worksheets("Ribbon").Range("A1")
  targetBook.Sheets("Notice").Columns("A:B").Copy ThisWorkbook.Worksheets("Style").Range("A1")
  targetBook.Sheets("Style").Columns("A:J").Copy ThisWorkbook.Worksheets("Style").Range("A1")

  Call Library.showNotice(200, "データコピー")


End Function


