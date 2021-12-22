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
        .maxCount0 = Selection.count
        .maxCount1 = Selection.count
        .maxCount2 = Selection.count
        .maxCount3 = Selection.count
        .maxCount4 = Selection.count
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
Function パターン選択()
  Dim line As Long, endLine As Long, count As Long
  Dim varDic
  Const funcName As String = "Library.パターン選択"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("パターン選択")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  If sampleDataList Is Nothing Then
    End
  End If
  For Each varDic In sampleDataList
    Select Case sampleDataList.Item(varDic)
      Case "0.氏名(姓)"
        Call 名前_姓(maxCount)
      
      Case "1.氏名(名)"
        Call 名前_名(maxCount)
        
      Case "1.氏名(フルネーム)"
        Call 名前_フルネーム(maxCount)
      
      Case Else
    End Select
    ActiveCell.Offset(0, 1).Select
  Next
  
  '処理終了--------------------------------------
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
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  Call showFrm_sampleData("【数値】桁数固定")
  
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "G/標準"
    Cells(line + count, ActiveCell.Column) = Library.makeRandomDigits(BK_setVal("digits"))
  Next

  '処理終了--------------------------------------
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
Function 数値_範囲()
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.数値_範囲"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【数値】範囲指定")
  line = Selection(1).Row
  maxCount = BK_setVal("maxCount")
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(BK_setVal("minVal"), BK_setVal("maxVal"))
  Next

  '処理終了--------------------------------------
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
  Const funcName As String = "Ctl_SampleData.名前定義削除"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  If maxCount <= 1 Then
    Call showFrm_sampleData("【名前】姓")
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine)
  Next
  '処理終了--------------------------------------
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
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("【名前】名")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("D" & getLine)
  Next
  
  '処理終了--------------------------------------
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
Function 名前_フルネーム(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.名前_フルネーム"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("【名前】フルネーム")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine) & "　" & BK_sheetTestData.Range("D" & getLine)
  Next

  '処理終了--------------------------------------
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
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【日付】日")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row

  fstDate = BK_setVal("minVal")
  lstDate = BK_setVal("maxVal")
  
  For count = 0 To (maxCount - 1)
    Randomize
    Cells(line + count, ActiveCell.Column) = Format(Int((lstDate - fstDate + 1) * Rnd + fstDate), "yyyy/mm/dd")
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd"
  Next
  
  '処理終了--------------------------------------
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
  
'  On Error GoTo catchError
  Call Library.startScript
  
  Call init.setting
  
  If maxCount <= 1 Then
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
  Next
  
  Call Library.endScript
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
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【日付】日")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row

  fstDate = BK_setVal("minVal")
  lstDate = BK_setVal("maxVal")
  
  line = ActiveCell.Row
  For count = 0 To maxCount - 1
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("09:00:00") * 100000, TimeValue("18:00:00") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = val
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  Next
  
  '処理終了--------------------------------------
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
  Const funcName As String = "Ctl_SampleData.その他_半角文字"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm("  " & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("【その他】文字")
  
  makeStr = ""
  If BK_setVal("strType01") Then makeStr = makeStr & SymbolCharacters
  If BK_setVal("strType02") Then makeStr = makeStr & HalfWidthCharacters
  If BK_setVal("strType03") Then makeStr = makeStr & StrConv(HalfWidthCharacters, vbLowerCase)
  If BK_setVal("strType04") Then makeStr = makeStr & HalfWidthDigit
  If BK_setVal("strType05") Then makeStr = makeStr & JapaneseCharacters
  If BK_setVal("strType06") Then makeStr = makeStr & StrConv(JapaneseCharacters, vbKatakana)
  If BK_setVal("strType07") Then makeStr = makeStr & MachineDependentCharacters
  
  For Each slctRange In Selection
    slctRange.Value = Library.makeRandomString(BK_setVal("maxCount"), makeStr)
  Next
  
  '処理終了--------------------------------------
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

