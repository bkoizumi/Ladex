Attribute VB_Name = "sampleData"
Dim newBook As Workbook
Dim count As Long, getLine As Long
Dim fstDate As Date, lstDate As Date

Public maxCount  As Long

'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showFrm_sampleData(showType As String)
  
  topPosition = Library.getRegistry("UserForm", "mkSmpDtTop")
  leftPosition = Library.getRegistry("UserForm", "mkSmpDtLeft")
  With Frm_smplData
    .StartUpPosition = 0
    If topPosition = "" Then
      .Top = 10
      .Left = 120
    Else
      .Top = topPosition
      .Left = leftPosition
    End If
    .Caption = showType
    
    '各ページ、パーツの有効/無効切り替え
    Select Case showType
      Case "パターン選択"
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
      
      Case "【数値】桁数固定"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame1.Caption = showType
      
      Case "【数値】範囲指定"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame2.Caption = showType
      
      Case "【名前】姓", "【名前】名", "【名前】フルネーム"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame3.Caption = showType
        
      Case "【日付】日", "【日付】時間", "【日付】日時"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        
        .Frame4.Caption = showType
      Case Else
    End Select
    
    .Show
  End With
  
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function パターン選択()
  Dim line As Long, endLine As Long
  Dim count As Long

'  On Error GoTo catchError
  Call init.setting
  
'  Sheets("Sheet1").Select
'  Sheets("Sheet1").Columns("A:Z").Clear
'  Sheets("Sheet1").Range("A1").Select
  
  Call showFrm_sampleData("パターン選択")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  If sampleDataList Is Nothing Then
    End
  End If
  For Each varDic In sampleDataList
    Debug.Print sampleDataList.Item(varDic)
    
    Select Case sampleDataList.Item(varDic)
      Case "0.氏名(姓)"
        Call 名前_姓(maxCount)
      
      Case "1.氏名(名)"
        Call 名前_名(maxCount)
        
      Case "1.氏名(フルネーム)"
        Call 名前_フルネーム(maxCount)
        
        
      Case Else
    End Select
    '次のセルに移動
    ActiveCell.Offset(0, 1).Select
    
  Next
    
  
  
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 数値_桁数固定()
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call init.setting
'  Sheets("Sheet1").Columns("A:A").Clear
'  Sheets("Sheet1").Range("A1").Select
  
  Call showFrm_sampleData("【数値】桁数固定")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column) = Library.makeRandomDigits(BK_setVal("digits"))
  Next


  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 数値_範囲()
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call init.setting
  
'  Sheets("Sheet1").Columns("B:B").Clear
'  Sheets("Sheet1").Range("B1").Select
  
  
  Call showFrm_sampleData("【数値】範囲指定")
  line = Selection(1).Row
  maxCount = BK_setVal("maxCount")
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(BK_setVal("minVal"), BK_setVal("maxVal"))
  Next


  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 名前_姓(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("C:C").Clear
'  Sheets("Sheet1").Range("C1").Select
  
  If maxCount = 0 Then
    Call showFrm_sampleData("【名前】姓")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine)
  Next

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 名前_名(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("D:D").Clear
'  Sheets("Sheet1").Range("D1").Select
  
  If maxCount = 0 Then
    Call showFrm_sampleData("【名前】名")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("D" & getLine)
  Next

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 名前_フルネーム(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("E:E").Clear
'  Sheets("Sheet1").Range("E1").Select
  
  If maxCount = 0 Then
    Call showFrm_sampleData("【名前】フルネーム")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine) & "　" & BK_sheetTestData.Range("D" & getLine)
  Next

  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 日付_日()
  Dim line As Long, endLine As Long
  
  
'  On Error GoTo catchError
  
  Call init.setting
  
  Sheets("Sheet1").Columns("F:F").Clear
  Sheets("Sheet1").Range("F1").Select
  
  Call showFrm_sampleData("【日付】日")
  line = Selection(1).Row
  'maxCount = BK_setVal("maxCount")
    
'  fstDate = BK_setVal("minVal")
'  lstDate = BK_setVal("maxVal")
  
  maxCount = 20
  fstDate = #4/1/2021#
  lstDate = #5/1/2020#
  
  For count = 0 To (maxCount - 1)
    Randomize
    Cells(line + count, ActiveCell.Column) = Format(Int((lstDate - fstDate + 1) * Rnd + fstDate), "yyyy/mm/dd")
    
  Next
  
  Exit Function
'エラー発生時====================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function 日付_時間()
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call init.setting
  
  Sheets("Sheet1").Columns("G:G").Clear
  Sheets("Sheet1").Range("G1").Select
  
  Call showFrm_sampleData("【日付】時間")
  line = Selection(1).Row
  maxCount = BK_setVal("maxCount")
    
'  fstDate = BK_setVal("minVal")
'  lstDate = BK_setVal("maxVal")
  
  For count = 0 To (maxCount - 1)
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("09:00:00") * 100000, TimeValue("18:00:00") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column) = Format(val, "hh:nn:ss")
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























