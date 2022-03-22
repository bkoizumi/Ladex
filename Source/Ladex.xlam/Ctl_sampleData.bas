Attribute VB_Name = "Ctl_sampleData"
Option Explicit

Dim newBook As Workbook
Dim count As Long, getLine As Long
Dim fstDate As Date, lstDate As Date

Public maxCount  As Long

'**************************************************************************************************
' * データ生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showFrm_sampleData(showType As String)
  With Frm_smplData
    .Caption = showType
    
    '各ページ、パーツの有効/無効切り替え
    Select Case showType
      Case "パターン選択"
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
      
      Case "【数値】桁数固定"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame1.Caption = showType
      
      Case "【数値】範囲指定"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame2.Caption = showType
      
      Case "【名前】姓", "【名前】名", "【名前】フルネーム"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame3.Caption = showType
        
      Case "【日付】日", "【日付】時間", "【日付】日時"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .minVal4 = #4/1/2021#
        .maxVal4 = #3/31/2022#
        
        .Frame4.Caption = showType
        
      Case "【その他】文字"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .maxCount5 = 25
        .strType01 = False
        .strType02 = False
        .strType03 = False
        .strType04 = False
        .strType05 = False
        .strType06 = False
        .strType07 = False
        
        .Frame5.Caption = showType
      
      Case Else
    End Select
    If Selection.count > 1 Then
      .maxCount0 = Selection.Rows.count
      .maxCount1 = Selection.Rows.count
      .maxCount2 = Selection.Rows.count
      .maxCount3 = Selection.Rows.count
      .maxCount4 = Selection.Rows.count
    End If

    .Show
  End With
  
  Exit Function

'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function パターン選択(Optional maxCount As Long)
  Dim line As Long, endLine As Long, count As Long, getLine As Long, getLine2 As Long
  Dim varDic
  Dim actRange As Range
  Dim strAddress As String
  Const funcName As String = "Ctl_SampleData.パターン選択"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("パターン選択")
  If sampleDataList Is Nothing Then
    End
  End If
  maxCount = BK_setVal("maxCount")

  Call Library.delSheetData(LadexSh_InputData)
  LadexSh_InputData.Cells.NumberFormatLocal = "@"
  
  line = 1
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row)
    getLine2 = Library.makeRandomNo(2, 5)
    
    '氏名(姓)
    LadexSh_InputData.Range("A" & line + count) = LadexSh_TestData.Range("A" & getLine)
    LadexSh_InputData.Range("D" & line + count) = LadexSh_TestData.Range("B" & getLine)
    
    '氏名(名)
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 4).End(xlUp).Row)
    LadexSh_InputData.Range("B" & line + count) = LadexSh_TestData.Range("D" & getLine)
    LadexSh_InputData.Range("E" & line + count) = LadexSh_TestData.Range("E" & getLine)
    
    LadexSh_InputData.Range("C" & line + count) = LadexSh_InputData.Range("A" & line + count) & "　" & LadexSh_InputData.Range("B" & line + count)
    LadexSh_InputData.Range("F" & line + count) = LadexSh_InputData.Range("D" & line + count) & "　" & LadexSh_InputData.Range("E" & line + count)
    
    '性別
    LadexSh_InputData.Range("G" & line + count) = LadexSh_TestData.Range("F" & getLine2)
    

    '血液型
    LadexSh_InputData.Range("H" & line + count) = LadexSh_TestData.Range("H" & getLine2)
    
    '生年月日
    LadexSh_InputData.Range("I" & line + count) = Format(Int((Date - #1/1/1950# + 1) * Rnd + #1/1/1950#), "yyyy/mm/dd")
    
    '年齢
    LadexSh_InputData.Range("J" & line + count) = Application.Evaluate("DATEDIF(""" & LadexSh_InputData.Range("I" & line + count) & """, TODAY(), ""Y"")")
    
    '電話番号
    LadexSh_InputData.Range("K" & line + count) = LadexSh_TestData.Range("Z" & getLine) & "-" & LadexSh_TestData.Range("AA" & getLine) & "-1234"
    
    'メールアドレス
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 10).End(xlUp).Row)
    LadexSh_InputData.Range("L" & line + count) = "Sample" & LadexSh_TestData.Range("J" & getLine)
    
    '都道府県
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 15).End(xlUp).Row)
    LadexSh_InputData.Range("M" & line + count) = LadexSh_TestData.Range("O" & getLine)
    LadexSh_InputData.Range("N" & line + count) = LadexSh_TestData.Range("P" & getLine)
    LadexSh_InputData.Range("O" & line + count) = LadexSh_TestData.Range("Q" & getLine)
    LadexSh_InputData.Range("P" & line + count) = LadexSh_TestData.Range("R" & getLine)
    LadexSh_InputData.Range("Q" & line + count) = LadexSh_TestData.Range("S" & getLine)
   
    If InStr(LadexSh_TestData.Range("U" & getLine2), "番") > 0 Then
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & LadexSh_TestData.Range("T" & getLine) & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbWide)
    Else
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & StrConv(Replace(LadexSh_TestData.Range("T" & getLine), "丁目", "-"), vbNarrow)
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbNarrow)
    End If
    
    LadexSh_InputData.Range("S" & line + count) = LadexSh_TestData.Range("V" & getLine)
    LadexSh_InputData.Range("T" & line + count) = LadexSh_TestData.Range("W" & getLine)
    LadexSh_InputData.Range("U" & line + count) = LadexSh_TestData.Range("X" & getLine)
    
    strAddress = LadexSh_InputData.Range("R" & line + count)
    strAddress = StrConv(Replace(strAddress, "丁目", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "丁", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "番地", ""), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "番", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "号", ""), vbNarrow)
    
    
    LadexSh_InputData.Range("V" & line + count) = strAddress
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  Set actRange = Selection(1)
  actRange.Select
  
  For Each varDic In sampleDataList
    Call Library.showDebugForm("varDic", varDic, "debug")
    Select Case varDic
      Case "氏名(姓)"
        LadexSh_InputData.Range("A1:A" & maxCount).Copy Selection
      Case "氏名(名)"
        LadexSh_InputData.Range("B1:B" & maxCount).Copy Selection

      Case "氏名(フルネーム)"
        LadexSh_InputData.Range("C1:C" & maxCount).Copy Selection

      Case "[カナ]氏名(姓)"
        LadexSh_InputData.Range("D1:D" & maxCount).Copy Selection
      Case "[カナ]氏名(名)"
        LadexSh_InputData.Range("E1:E" & maxCount).Copy Selection
      Case "[カナ]氏名(フルネーム)"
        LadexSh_InputData.Range("F1:F" & maxCount).Copy Selection
      Case "性別"
        LadexSh_InputData.Range("G1:G" & maxCount).Copy Selection
      Case "血液型"
        LadexSh_InputData.Range("H1:H" & maxCount).Copy Selection
      Case "生年月日"
        LadexSh_InputData.Range("I1:I" & maxCount).Copy Selection
      Case "年齢"
        LadexSh_InputData.Range("J1:J" & maxCount).Copy Selection
      Case "電話番号"
        LadexSh_InputData.Range("K1:K" & maxCount).Copy Selection
      Case "メールアドレス"
        LadexSh_InputData.Range("L1:L" & maxCount).Copy Selection
      Case "都道府県コード"
        LadexSh_InputData.Range("M1:M" & maxCount).Copy Selection
      Case "郵便番号"
        LadexSh_InputData.Range("N1:N" & maxCount).Copy Selection
      Case "都道府県"
        LadexSh_InputData.Range("O1:O" & maxCount).Copy Selection
      Case "市区郡町村"
        LadexSh_InputData.Range("P1:P" & maxCount).Copy Selection
      Case "町域"
        LadexSh_InputData.Range("Q1:Q" & maxCount).Copy Selection
      Case "丁目・字名・番地"
        LadexSh_InputData.Range("R1:R" & maxCount).Copy Selection
      Case "[カナ]都道府県"
        LadexSh_InputData.Range("S1:S" & maxCount).Copy Selection
      Case "[カナ]市区郡町村"
        LadexSh_InputData.Range("T1:T" & maxCount).Copy Selection
      Case "[カナ]町域"
        LadexSh_InputData.Range("U1:U" & maxCount).Copy Selection
      Case "[カナ]丁目・字名・番地"
        LadexSh_InputData.Range("V1:V" & maxCount).Copy Selection
      
      Case Else
    End Select
    ActiveCell.Offset(0, 1).Select
    DoEvents
  Next
  actRange.Select
  Call Library.delSheetData(LadexSh_InputData)
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
    Call Library.errorHandle
End Function

'==================================================================================================
Function 数値_桁数固定(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前定義削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  Call showFrm_sampleData("【数値】桁数固定")
  
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "###"
    Cells(line + count, ActiveCell.Column) = BK_setVal("addFirst") & Library.makeRandomDigits(BK_setVal("digits")) & BK_setVal("addEnd")
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 数値_範囲(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.数値_範囲"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【数値】範囲指定")
  line = Selection(1).Row
  
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "###"
    Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(BK_setVal("minVal"), BK_setVal("maxVal"))
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 名前_姓(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前_姓"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If IsMissing(maxCount) Then
    Call showFrm_sampleData("【名前】姓")
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("A" & getLine)
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 名前_名(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前_名"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If IsMissing(maxCount) Then
    Call showFrm_sampleData("【名前】名")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("D" & getLine)
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 名前_フルネーム(Optional maxCount As Long, Optional kanaFlg As Boolean = False)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前_フルネーム"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If IsMissing(maxCount) Then
    Call showFrm_sampleData("【名前】フルネーム")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("A" & getLine) & "　" & LadexSh_TestData.Range("D" & getLine)
    
    If kanaFlg = True Then
      Cells(line + count, ActiveCell.Column + 1) = LadexSh_TestData.Range("B" & getLine) & "　" & LadexSh_TestData.Range("E" & getLine)
    End If
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 日付_日(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前_フルネーム"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  Call showFrm_sampleData("【日付】日")
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row

  fstDate = BK_setVal("minVal")
  lstDate = BK_setVal("maxVal")
  
  For count = 0 To (maxCount - 1)
    Randomize
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd"
    Cells(line + count, ActiveCell.Column) = Format(Int((lstDate - fstDate + 1) * Rnd + fstDate), "yyyy/mm/dd")
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 日付_時間(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  Const funcName As String = "Ctl_SampleData.日付_時間"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Ctl_ProgressBar.showStart
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  If IsMissing(maxCount) Then
    Call showFrm_sampleData("【日付】時間")
    maxCount = BK_setVal("maxCount")
  End If
  
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("00:00:00") * 100000, TimeValue("23:59:59") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = Format(val, "hh:nn:ss")
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "hh:mm:ss"
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 日時(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  Const funcName As String = "Ctl_SampleData.日付_時間"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Ctl_ProgressBar.showStart
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【日付】日")
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row

  fstDate = BK_setVal("minVal")
  lstDate = BK_setVal("maxVal")
  
  line = Selection(1).Row
  For count = 0 To maxCount - 1
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("09:00:00") * 100000, TimeValue("18:00:00") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = val
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function その他_文字(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim makeStr As String
  Dim slctRange As Range
  Const funcName As String = "Ctl_SampleData.その他_文字"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【その他】文字")
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  
  makeStr = ""
  If BK_setVal("strType01") Then makeStr = makeStr & SymbolCharacters
  If BK_setVal("strType02") Then makeStr = makeStr & HalfWidthCharacters
  If BK_setVal("strType03") Then makeStr = makeStr & StrConv(HalfWidthCharacters, vbLowerCase)
  If BK_setVal("strType04") Then makeStr = makeStr & HalfWidthDigit
  If BK_setVal("strType05") Then makeStr = makeStr & JapaneseCharacters
  If BK_setVal("strType06") Then makeStr = makeStr & StrConv(JapaneseCharacters, vbKatakana)
  If BK_setVal("strType07") Then makeStr = makeStr & MachineDependentCharacters

  For Each slctRange In Selection
    slctRange.Value = Library.makeRandomString(maxCount, makeStr)
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function メールアドレス(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim makeStr As String
  Const funcName As String = "Ctl_SampleData.メールアドレス"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 10).End(xlUp).Row
  If IsMissing(maxCount) Then
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    
    makeStr = ""
    makeStr = makeStr & HalfWidthCharacters
    makeStr = makeStr & StrConv(HalfWidthCharacters, vbLowerCase)
    makeStr = makeStr & HalfWidthDigit
    makeStr = Library.makeRandomString(10, makeStr)
    
    Cells(line + count, ActiveCell.Column) = "Sample." & makeStr & LadexSh_TestData.Range("J" & getLine)
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 住所(maxCount As Long, ParamArray addressFlgs())
  Dim line As Long, endLine As Long
  Dim getLine As Long, getLine2 As Long
  Dim strAddress As String
  Const funcName As String = "Ctl_SampleData.住所"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 16).End(xlUp).Row
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    getLine2 = Library.makeRandomNo(2, 5)
    
    If InStr(LadexSh_TestData.Range("U" & getLine2), "番") > 0 Then
      strAddress = LadexSh_TestData.Range("P" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("Q" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("R" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("S" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("T" & getLine) & vbTab
      strAddress = strAddress & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbWide)
    Else
      strAddress = LadexSh_TestData.Range("P" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("Q" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("R" & getLine) & vbTab
      strAddress = strAddress & LadexSh_TestData.Range("S" & getLine) & vbTab
      strAddress = strAddress & StrConv(Replace(LadexSh_TestData.Range("T" & getLine), "丁目", "-"), vbUpperCase)
      strAddress = strAddress & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbNarrow)
    End If
    
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = strAddress
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 電話番号(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.電話番号"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 15).End(xlUp).Row
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("Y" & getLine) & "-" & LadexSh_TestData.Range("Z" & getLine) & "-1234"
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "データ生成")
  Next
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

