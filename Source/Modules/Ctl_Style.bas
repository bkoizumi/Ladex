Attribute VB_Name = "Ctl_Style"
Option Explicit

Dim setStyleBook     As Workbook


'**************************************************************************************************
' * �X�^�C��Import/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Export()
  Dim filePath As String, fileName As String
  Dim FSO As Object
     
     
  '�����J�n--------------------------------------
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


  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
Function Import()
  Dim FSO As Object
  Dim styleBookPath As String
  Dim filePath As String, fileName As String
  
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
     
     
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.Import"

  Call Library.startScript
  Call init.setting
  
  '----------------------------------------------
  If setStyleBook Is Nothing Then
    Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
  End If
  
  Call Library.startScript
  setStyleBook.Save
  
  setStyleBook.Sheets("Style").Columns("A:J").Copy BK_ThisBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  styleBookPath = setStyleBook.Path & "\" & setStyleBook.Name
  Application.DisplayAlerts = False
  setStyleBook.Close
  Call Library.execDel(styleBookPath)
  
  Set setStyleBook = Nothing
  If MsgBox("�X�^�C����K�����܂����H", vbYesNo + vbExclamation) = vbYes Then
    Call Ctl_Style.�X�^�C���폜
    Call Ctl_Style.�X�^�C���ݒ�
  End If
  
  
  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'**************************************************************************************************
' * �X�^�C���폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �X�^�C���폜()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
  On Error Resume Next
  
  
  Call Library.startScript
  Call init.setting
  
  '�u�b�N�̕ی�m�F
  If ActiveWorkbook.ProtectWindows = True Then
    Call Library.showNotice(412, , True)
  End If

  '�V�[�g�̕ی�m�F
  For Each tempSheet In Sheets
    If Worksheets(tempSheet.Name).ProtectContents = True Then
      Worksheets(tempSheet.Name).Select
      Call Library.showNotice(413, , True)
    End If
  Next
  
  
  count = 1
  Call Ctl_ProgressBar.showStart
  endCount = ActiveWorkbook.Styles.count
  
  For Each s In ActiveWorkbook.Styles
    Call Ctl_ProgressBar.showCount("��`�σX�^�C���폜", count, endCount, s.Name)
    Select Case s.Name
      Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Call Library.showDebugForm("��`�σX�^�C��    �F" & s.Name)
      Case Else
        Call Library.showDebugForm("��`�σX�^�C���폜�F" & s.Name)
        s.delete
    End Select
    count = count + 1
  Next
  
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript

End Function


'**************************************************************************************************
' * �X�^�C���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �X�^�C���ݒ�()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
  On Error Resume Next
  
  
  Call Library.startScript
  Call init.setting
  
  Call Ctl_Style.�X�^�C���폜
  Call Ctl_ProgressBar.showStart

  
  '�X�^�C��������----------------------------------------------------------------------------------
  endLine = BK_sheetStyle.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetStyle.Range("A" & line) <> "����" Then
      Call Ctl_ProgressBar.showCount("�X�^�C��������", line, endLine, BK_sheetStyle.Range("B" & line))
      Call Library.showDebugForm("�X�^�C���������F" & BK_sheetStyle.Range("B" & line))

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

        '�z�u
        If BK_sheetStyle.Range("F" & line) = "TRUE" Then
          .HorizontalAlignment = BK_sheetStyle.Range("J" & line).HorizontalAlignment
          .VerticalAlignment = BK_sheetStyle.Range("J" & line).VerticalAlignment
        End If

        '�r��
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


        '�w�i�F
        If BK_sheetStyle.Range("H" & line) = "TRUE" Then
          .Interior.Color = BK_sheetStyle.Range("J" & line).Interior.Color
        End If


      End With
    End If
  Next
  
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript

End Function


'==================================================================================================
Function �X�^�C��������()
  Dim FSO As Object
  Dim setActivBook     As Workbook
  Dim filePath As String, fileName As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.�X�^�C��������"

  Call Library.startScript
  Call init.setting
  Call Library.showDebugForm(FuncName & "�J�n==========================================")
  '----------------------------------------------
  Call Ctl_Style.�X�^�C���폜

  Set setActivBook = ActiveWorkbook
  Set setStyleBook = Workbooks.add
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  With setStyleBook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      filePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs filePath
  End With
  
  setActivBook.Activate
  ActiveWorkbook.Styles.Merge Workbook:=Workbooks(fileName)
  Set FSO = Nothing
  setStyleBook.Close
  
  Call Library.execDel(filePath)
  
  '�����I��--------------------------------------
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm(FuncName & "�I��==========================================")
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function
