Attribute VB_Name = "sampleData"
Dim newBook As Workbook
Dim count As Long, getLine As Long
Dim fstDate As Date, lstDate As Date

Public Const maxCount  As Long = 132

Function ExcelHelp142()
  
  Call Library.delSheetData(2)
  Call sampleData.乱数
  Call sampleData.氏名
  Call sampleData.帳票名
  Call sampleData.フェーズ
  Call sampleData.結果
  Call sampleData.日時

  Range("B2:B22").Copy
  For line = 23 To maxCount Step 20
    Range("B" & line).Select
    ActiveSheet.Paste
  Next
  Application.CutCopyMode = False
  Range("A2:G" & maxCount + 1).Select

End Function


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
    endLine = sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
    
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
    endLine = sheetTestData.Cells(Rows.count, 4).End(xlUp).Row
    
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
    endLine = sheetTestData.Cells(Rows.count, 8).End(xlUp).Row
    
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
    endLine = sheetTestData.Cells(Rows.count, 11).End(xlUp).Row
    
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


'==================================================================================================
Function 乱数()
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call init.setting
  Range("A2").Select
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    Cells(line + count, ActiveCell.Column) = 11111 & Library.makeRandomDigits(6)
  Next


  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 氏名()
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
  Range("B2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine) & "　" & BK_sheetTestData.Range("D" & getLine)
  Next

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 日付()
  Dim line As Long, endLine As Long
  
  
'  On Error GoTo catchError
  
  Call init.setting
  
  fstDate = #4/1/2011# 'ランダムに作成する最初の日
  lstDate = #9/30/2011# 'ランダムに作成する最後の日

  Range("F2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    Randomize
    Cells(line + count, ActiveCell.Column) = Int((lstDate - fstDate + 1) * Rnd + fstDate)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 日時()
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call init.setting
  
  fstDate = DateAdd("d", -10, Date)
  lstDate = Date
  
  Range("F2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("09:00:00") * 100000, TimeValue("18:00:00") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column) = val
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    
    Cells(line + count, ActiveCell.Column + 1) = DateAdd("s", Library.makeRandomNo(0, 600), val)
    Cells(line + count, ActiveCell.Column + 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 帳票名()
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call init.setting
  
  targetString = Array("源泉徴収票", "給与明細票", "確定申告書", "納税・課税通知書", "納税証明書", "所得証明書", "青色申告書", "収支内訳書", "支払調書", "年金証書", "年金通知書")

  
  Range("C2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    getRand = Library.makeRandomNo(0, 10)
    
    Cells(line + count, ActiveCell.Column) = targetString(getRand)
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function フェーズ()
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call init.setting
  
'  targetString = Array("エスカレ", "1_1", "1_2", "1_3", "1_4", "2_1", "2_2", "2_3", "2_4")
  targetString = Array("1_1", "2_1")

  
  Range("D2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    getRand = Library.makeRandomNo(0, 1)
    
    Cells(line + count, ActiveCell.Column) = targetString(getRand)
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 結果()
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call init.setting
  
'  targetString = Array("ｴｽｶﾚ依頼", "ｴｽｶﾚ済", "確認完了", "差し戻し")
  targetString = Array("確認完了")

  
  Range("E2").Select
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    getRand = Library.makeRandomNo(0, 0)
    
    Cells(line + count, ActiveCell.Column) = targetString(getRand)
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


















