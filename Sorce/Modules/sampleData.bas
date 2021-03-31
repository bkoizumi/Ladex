Attribute VB_Name = "sampleData"

'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function サンプルデータ()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim getLine As Long, count As Long, getMaxCount As Long
  
  Dim newBook As Workbook
'  On Error GoTo catchError
  setName01 = True   '姓
  setName02 = True    '名
  setName = True      '姓名
  setPref = True      '都道府県
  setMail = True      'メールアドレス
  
  
  Call init.setting
  Set newBook = Workbooks.add
  ThisWorkbook.Activate
  sheetTestData.Activate
  
  
  getMaxCount = 100
  If setName01 = True Then
    colLine = 1
    endLine = sheetTestData.Cells(Rows.count, 1).End(xlUp).row
    
    newBook.Sheets("Sheet1").Cells(1, colLine) = "姓"
    newBook.Sheets("Sheet1").Cells(1, colLine + 1) = "セイ"
    For count = 1 To getMaxCount
      getLine = Library.makeRandomNo(2, endLine)
      
      newBook.Sheets("Sheet1").Cells(count + 1, colLine) = sheetTestData.Range("A" & getLine)
      newBook.Sheets("Sheet1").Cells(count + 1, colLine + 1) = sheetTestData.Range("B" & getLine)
    Next
  End If
  
  If setName02 = True Then
    colLine = newBook.Sheets("Sheet1").Cells(1, Columns.count).End(xlToLeft).Column + 1
    endLine = sheetTestData.Cells(Rows.count, 4).End(xlUp).row
    
    newBook.Sheets("Sheet1").Cells(1, colLine) = "名"
    newBook.Sheets("Sheet1").Cells(1, colLine + 1) = "メイ"
    newBook.Sheets("Sheet1").Cells(1, colLine + 2) = "性別"
    
    For count = 1 To getMaxCount
      getLine = Library.makeRandomNo(2, endLine)
      
      newBook.Sheets("Sheet1").Cells(count + 1, colLine) = sheetTestData.Range("D" & getLine)
      newBook.Sheets("Sheet1").Cells(count + 1, colLine + 1) = sheetTestData.Range("E" & getLine)
      newBook.Sheets("Sheet1").Cells(count + 1, colLine + 2) = sheetTestData.Range("F" & getLine)
    Next
  End If
  
  If setName = True Then
    colLine = newBook.Sheets("Sheet1").Cells(1, Columns.count).End(xlToLeft).Column + 1
    
    newBook.Sheets("Sheet1").Cells(1, colLine) = "姓名"
    newBook.Sheets("Sheet1").Cells(1, colLine + 1) = "セイメイ"
    For count = 1 To getMaxCount
      newBook.Sheets("Sheet1").Cells(count + 1, colLine) = newBook.Sheets("Sheet1").Range("A" & count + 1) & " " & newBook.Sheets("Sheet1").Range("C" & count + 1)
      newBook.Sheets("Sheet1").Cells(count + 1, colLine + 1) = newBook.Sheets("Sheet1").Range("B" & count + 1) & " " & newBook.Sheets("Sheet1").Range("D" & count + 1)
    Next
  End If
  
  
  If setPref = True Then
    colLine = newBook.Sheets("Sheet1").Cells(1, Columns.count).End(xlToLeft).Column + 1
    endLine = sheetTestData.Cells(Rows.count, 8).End(xlUp).row
    
    newBook.Sheets("Sheet1").Cells(1, colLine) = "都道府県"
    newBook.Sheets("Sheet1").Cells(1, colLine + 1) = "都道府県コード"
    For count = 1 To getMaxCount
      getLine = Library.makeRandomNo(2, endLine)
      
      newBook.Sheets("Sheet1").Cells(count + 1, colLine) = sheetTestData.Range("I" & getLine)
      newBook.Sheets("Sheet1").Cells(count + 1, colLine + 1) = sheetTestData.Range("H" & getLine)
    Next
  End If
    
   If setMail = True Then
    colLine = newBook.Sheets("Sheet1").Cells(1, Columns.count).End(xlToLeft).Column + 1
    endLine = sheetTestData.Cells(Rows.count, 11).End(xlUp).row
    
    newBook.Sheets("Sheet1").Cells(1, colLine) = "メールアドレス"
    For count = 1 To getMaxCount
      getLine = Library.makeRandomNo(2, endLine)
      
      newBook.Sheets("Sheet1").Cells(count + 1, colLine) = "testMail" & sheetTestData.Range("K" & getLine)
    Next
  End If
  
  
  
  
  
  
  
  

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
