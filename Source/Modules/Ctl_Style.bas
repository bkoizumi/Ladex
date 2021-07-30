Attribute VB_Name = "Ctl_Style"
Option Explicit

Dim setStyleBook     As Workbook


'**************************************************************************************************
' * スタイルImport/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Export()
  Dim filePath As String, fileName As String
  Dim FSO As Object
     
     
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.Export"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------

  BK_sheetStyle.Copy
  
  Set setStyleBook = ActiveWorkbook
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  With setStyleBook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      filePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs filePath
  End With
  Set FSO = Nothing
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", filePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'==================================================================================================
Function Import()
  Dim FSO As Object
  Dim filePath As String, fileName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
     
     
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.Import"

  Call Library.startScript
  Call init.setting
  
  '----------------------------------------------
  If setStyleBook Is Nothing Then
    Call Library.showNotice(400, FuncName, True)
  End If
  
  Call Library.startScript
  setStyleBook.Save
  
  setStyleBook.Sheets("Style").Columns("A:J").Copy BK_ThisBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  Application.DisplayAlerts = False
  setStyleBook.Close
  Kill setStyleBook.Path
  
  Set setStyleBook = Nothing
  Call Ctl_Style.スタイル削除
  
  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'**************************************************************************************************
' * スタイル削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function スタイル削除()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  
'  On Error Resume Next
  
  Call Library.startScript
  Call init.setting
  
  count = 1
  Call Ctl_ProgressBar.showStart
  endCount = ActiveWorkbook.Styles.count
  
  For Each s In ActiveWorkbook.Styles
    Call Ctl_ProgressBar.showCount("定義済スタイル削除", count, endCount, s.Name)
    Call Library.showDebugForm("定義済スタイル削除：" & s.Name)
    
    Select Case s.Name
      Case "Normal"
      Case Else
        s.delete
    End Select
    count = count + 1
  Next
  
  'スタイル初期化----------------------------------------------------------------------------------
  endLine = BK_sheetStyle.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetStyle.Range("A" & line) <> "無効" Then
      Call Ctl_ProgressBar.showCount("スタイル初期化", line, endLine, BK_sheetStyle.Range("B" & line))
      Call Library.showDebugForm("スタイル初期化：" & BK_sheetStyle.Range("B" & line))
      
      If BK_sheetStyle.Range("B" & line) <> "Normal" Then
        ActiveWorkbook.Styles.add Name:=BK_sheetStyle.Range("B" & line).Value
      End If
      
      With ActiveWorkbook.Styles(BK_sheetStyle.Range("B" & line).Value)
        
        If BK_sheetStyle.Range("C" & line) <> "" Then
          .NumberFormatLocal = BK_sheetStyle.Range("C" & line)
        End If
        
        .IncludeNumber = BK_sheetStyle.Range("D" & line)
        .IncludeFont = BK_sheetStyle.Range("E" & line)
        .IncludeAlignment = BK_sheetStyle.Range("F" & line)
        .IncludeBorder = BK_sheetStyle.Range("G" & line)
        .IncludePatterns = BK_sheetStyle.Range("H" & line)
        .IncludeProtection = BK_sheetStyle.Range("I" & line)
        
        If BK_sheetStyle.Range("E" & line) = "TRUE" Then
          .Font.Name = BK_sheetStyle.Range("J" & line).Font.Name
          .Font.Size = BK_sheetStyle.Range("J" & line).Font.Size
          .Font.Color = BK_sheetStyle.Range("J" & line).Font.Color
          .Font.Bold = BK_sheetStyle.Range("J" & line).Font.Bold
        End If
        
        '配置
        If BK_sheetStyle.Range("F" & line) = "TRUE" Then
          .HorizontalAlignment = BK_sheetStyle.Range("J" & line).HorizontalAlignment
          .VerticalAlignment = BK_sheetStyle.Range("J" & line).VerticalAlignment
        End If
        
        '罫線
        If BK_sheetStyle.Range("G" & line) = "TRUE" Then
          If BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).LineStyle <> xlNone Then
            .Borders(xlDiagonalDown).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).LineStyle
            .Borders(xlDiagonalDown).Weight = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).Weight
            .Borders(xlDiagonalDown).Color = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalDown).Color
          End If
          
          If BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).LineStyle <> xlNone Then
            .Borders(xlDiagonalUp).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).LineStyle
            .Borders(xlDiagonalUp).Weight = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).Weight
            .Borders(xlDiagonalUp).Color = BK_sheetStyle.Range("J" & line).Borders(xlDiagonalUp).Color
          End If
          
          If BK_sheetStyle.Range("J" & line).Borders(xlLeft).LineStyle <> xlNone Then
            .Borders(xlLeft).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlLeft).LineStyle
            .Borders(xlLeft).Weight = BK_sheetStyle.Range("J" & line).Borders(xlLeft).Weight
            .Borders(xlLeft).Color = BK_sheetStyle.Range("J" & line).Borders(xlLeft).Color
          End If
          
          If BK_sheetStyle.Range("J" & line).Borders(xlRight).LineStyle <> xlNone Then
            .Borders(xlRight).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlRight).LineStyle
            .Borders(xlRight).Weight = BK_sheetStyle.Range("J" & line).Borders(xlRight).Weight
            .Borders(xlRight).Color = BK_sheetStyle.Range("J" & line).Borders(xlRight).Color
          End If
          
          If BK_sheetStyle.Range("J" & line).Borders(xlTop).LineStyle <> xlNone Then
            .Borders(xlTop).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlTop).LineStyle
            .Borders(xlTop).Weight = BK_sheetStyle.Range("J" & line).Borders(xlTop).Weight
            .Borders(xlTop).Color = BK_sheetStyle.Range("J" & line).Borders(xlTop).Color
          End If
          
          If BK_sheetStyle.Range("J" & line).Borders(xlBottom).LineStyle <> xlNone Then
            .Borders(xlBottom).LineStyle = BK_sheetStyle.Range("J" & line).Borders(xlBottom).LineStyle
            .Borders(xlBottom).Weight = BK_sheetStyle.Range("J" & line).Borders(xlBottom).Weight
            .Borders(xlBottom).Color = BK_sheetStyle.Range("J" & line).Borders(xlBottom).Color
          End If
        End If
        
        
        '背景色
        If BK_sheetStyle.Range("H" & line) = "TRUE" Then
          .Interior.Color = BK_sheetStyle.Range("J" & line).Interior.Color
        End If
        
        
      End With
    End If
  Next
  
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript

End Function
