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
  Const funcName As String = "Ctl_Style.Export"
     
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  BK_sheetStyle.Copy
  
  Set setStyleBook = ActiveWorkbook
  setStyleBook.SaveAs LadexDir & "\" & "�X�^�C�����.xlsx"
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", filePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function Import()
  Dim styleBookPath As String
  Dim filePath As String, fileName As String
  Const funcName As String = "Ctl_Style.Import"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Ctl_ProgressBar.showStart
  PrgP_Max = 4
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  If Library.chkIsOpen("�X�^�C�����.xlsx") Then
    Set setStyleBook = Workbooks("�X�^�C�����.xlsx")
    setStyleBook.Save
  Else
    Set setStyleBook = Workbooks.Open(LadexDir & "\" & "�X�^�C�����.xlsx")
    Call Library.startScript
  End If
  setStyleBook.Sheets("Style").Columns("A:J").Copy BK_ThisBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  styleBookPath = setStyleBook.Path & "\" & setStyleBook.Name
  Application.DisplayAlerts = False
  setStyleBook.Close
'  Call Library.execDel(styleBookPath)
  
  Set setStyleBook = Nothing
  If MsgBox("�X�^�C����K�����܂����H", vbYesNo + vbExclamation) = vbYes Then
    Call Ctl_Style.�X�^�C���ݒ�
  End If
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
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
  Const funcName As String = "Ctl_Style.�X�^�C���폜"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
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
  endCount = ActiveWorkbook.Styles.count
  
  For Each s In ActiveWorkbook.Styles
    Call Ctl_ProgressBar.showCount("��`�σX�^�C���폜", 1, 2, count, endCount, s.Name)
    Select Case s.Name
      Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Call Library.showDebugForm("��`�σX�^�C��", s.Name, "debug")
      
      'Ladex�̏����ݒ�
      Case "����؂�", "�p�[�Z���g", "�ʉ�", "�ʉ�[��P��]", "���l", "���l[��P��]", "00.0", "���t [yyyy/mm/dd]", "���t [yyyy/m]", "����", "�s�v", "Error", "�v�m�F", "H_�W��", "H_�ڎ�1", "H_�ڎ�2", "H_�ڎ�3", "�s�t"
        Call Library.showDebugForm("Ladex�X�^�C�� ", s.Name, "debug")
      
      Case Else
        Call Library.showDebugForm("�폜�X�^�C��  ", s.Name, "debug")
        s.delete
    End Select
    count = count + 1
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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
  
'  On Error Resume Next
  Const funcName As String = "Ctl_Style.�X�^�C���ݒ�"
  
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call Ctl_Style.�X�^�C���폜
  
  '�X�^�C��������--------------------------------
  endLine = BK_sheetStyle.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetStyle.Range("A" & line) <> "����" Then
      Call Ctl_ProgressBar.showCount("�X�^�C���ݒ�", 1, 2, line, endLine, BK_sheetStyle.Range("B" & line))

      Select Case BK_sheetStyle.Range("B" & line)
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
          Call Library.showDebugForm("��`�σX�^�C��", BK_sheetStyle.Range("B" & line), "debug")
          
      'Ladex�̏����ݒ�
      Case "����؂�", "�p�[�Z���g", "�ʉ�", "�ʉ�[��P��]", "���l", "���l[��P��]", "00.0", "���t [yyyy/mm/dd]", "���t [yyyy/m]", "����", "�s�v", "Error", "�v�m�F", "H_�W��", "H_�ڎ�1", "H_�ڎ�2", "H_�ڎ�3", "�s�t"
        Call Library.showDebugForm("Ladex�X�^�C�� ", BK_sheetStyle.Range("B" & line), "debug")
      
      Case Else
        Call Library.showDebugForm("�X�^�C����", BK_sheetStyle.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=BK_sheetStyle.Range("B" & line).Value
      End Select

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
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
      Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �X�^�C��������()
  Dim FSO As Object
  Dim setActivBook     As Workbook
  Dim filePath As String, fileName As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Style.�X�^�C��������"

  Call Library.startScript
  Call init.setting
  Call Library.showDebugForm(funcName & "�J�n==========================================")
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
  Call Library.showDebugForm("", , "end")
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function
