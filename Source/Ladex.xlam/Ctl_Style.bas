Attribute VB_Name = "Ctl_Style"
Option Explicit

Dim setStyleBook     As Workbook


'**************************************************************************************************
' * �X�^�C��Import/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �X�^�C���o��()
  Dim FilePath As String, fileName As String
  Const funcName As String = "Ctl_Style.�X�^�C���o��"
     
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
  LadexSh_Style.copy
  
  Set setStyleBook = ActiveWorkbook
  setStyleBook.SaveAs LadexDir & "\" & "�X�^�C�����.xlsx"
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", FilePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������-------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function �X�^�C���捞()
  Dim styleBookPath As String
  Dim FilePath As String, fileName As String
  Const funcName As String = "Ctl_Style.�X�^�C���捞"
  
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
  setStyleBook.Sheets("Style").Columns("A:J").copy LadexBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")
  
  styleBookPath = setStyleBook.path & "\" & setStyleBook.Name
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
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������-------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


'**************************************************************************************************
' * �X�^�C���폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �X�^�C���폜()
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  Dim useStyleName As Variant
  
  Const funcName As String = "Ctl_Style.�X�^�C���폜"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("�X�^�C�����p�m�F")
  PrgP_Cnt = PrgP_Cnt + 1
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
  
  Call Library.�X�^�C�����p�m�F
  Call Library.�X�^�C���폜

  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �X�^�C��_�S�폜()
  Dim s
  Dim count As Long, endCount As Long
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  Dim useStyleName As Variant
  
  Const funcName As String = "Ctl_Style.�X�^�C��_�S�폜"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("�X�^�C�����p�m�F")
  PrgP_Cnt = PrgP_Cnt + 1
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
  
    
  ReDim useStyle(0)
  useStyle(0) = "�W��"
  Call Library.�X�^�C���폜

  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

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
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
  Const funcName As String = "Ctl_Style.�X�^�C���ݒ�"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("�X�^�C�����p�m�F")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  '�����X�^�C���폜------------------------------
  Call Library.�X�^�C�����p�m�F
  Call Library.�X�^�C���폜
  
  If setStyleBook Is Nothing Then
    If Library.chkIsOpen("�X�^�C�����.xlsx") Then
      Set setStyleBook = Workbooks("�X�^�C�����.xlsx")
      setStyleBook.Save
    Else
      Set setStyleBook = Workbooks.Open(LadexDir & "\" & "�X�^�C�����.xlsx")
      Call Library.startScript
    End If
    setStyleBook.Sheets("Style").Columns("A:J").copy LadexBook.Worksheets("Style").Range("A1")
    setStyleBook.Close
  End If
  Set setStyleBook = Nothing
  
  
  '�X�^�C��������--------------------------------
  endLine = LadexSh_Style.Cells(Rows.count, 2).End(xlUp).Row
  For line = 2 To endLine
    On Error Resume Next
    
    If LadexSh_Style.Range("A" & line) <> "����" Then
      Call Ctl_ProgressBar.showCount("�X�^�C���ݒ�", 1, 2, line, endLine, LadexSh_Style.Range("B" & line))

      Select Case LadexSh_Style.Range("B" & line)
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
          Call Library.showDebugForm("��`�σX�^�C��", LadexSh_Style.Range("B" & line), "debug")
          
      'Ladex�̏����ݒ�
      Case "����؂�", "�p�[�Z���g", "�ʉ�", "�ʉ�[��P��]", "���l", "���l[��P��]", "00.0", "���t [yyyy/mm/dd]", "���t [yyyy/m]", "����", "�s�v", "Error", "�v�m�F", "H_�W��", "H_�ڎ�1", "H_�ڎ�2", "H_�ڎ�3", "�s�t"
        Call Library.showDebugForm("Ladex�X�^�C�� ", LadexSh_Style.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=LadexSh_Style.Range("B" & line).Value
      Case Else
        Call Library.showDebugForm("�X�^�C����", LadexSh_Style.Range("B" & line), "debug")
        ActiveWorkbook.Styles.add Name:=LadexSh_Style.Range("B" & line).Value
      End Select

      With ActiveWorkbook.Styles(LadexSh_Style.Range("B" & line).Value)

        If LadexSh_Style.Range("C" & line) <> "" Then
          .NumberFormatLocal = LadexSh_Style.Range("C" & line)
        End If

        .IncludeNumber = LadexSh_Style.Range("D" & line)
        .IncludeFont = LadexSh_Style.Range("E" & line)
        .IncludeAlignment = LadexSh_Style.Range("F" & line)
        .IncludeBorder = LadexSh_Style.Range("G" & line)
        .IncludePatterns = LadexSh_Style.Range("H" & line)
        .IncludeProtection = LadexSh_Style.Range("I" & line)

        If LadexSh_Style.Range("E" & line) = "TRUE" Then
          .Font.Name = LadexSh_Style.Range("J" & line).Font.Name
          .Font.Size = LadexSh_Style.Range("J" & line).Font.Size
          .Font.Color = LadexSh_Style.Range("J" & line).Font.Color
          .Font.Bold = LadexSh_Style.Range("J" & line).Font.Bold
        End If

        '�z�u
        If LadexSh_Style.Range("F" & line) = "TRUE" Then
          .HorizontalAlignment = LadexSh_Style.Range("J" & line).HorizontalAlignment
          .VerticalAlignment = LadexSh_Style.Range("J" & line).VerticalAlignment
        End If

        '�r��
        If LadexSh_Style.Range("G" & line) = "TRUE" Then
          If LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).LineStyle <> xlNone Then
            .Borders(xlDiagonalDown).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).LineStyle
            .Borders(xlDiagonalDown).Weight = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).Weight
            .Borders(xlDiagonalDown).Color = LadexSh_Style.Range("J" & line).Borders(xlDiagonalDown).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).LineStyle <> xlNone Then
            .Borders(xlDiagonalUp).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).LineStyle
            .Borders(xlDiagonalUp).Weight = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).Weight
            .Borders(xlDiagonalUp).Color = LadexSh_Style.Range("J" & line).Borders(xlDiagonalUp).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlLeft).LineStyle <> xlNone Then
            .Borders(xlLeft).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlLeft).LineStyle
            .Borders(xlLeft).Weight = LadexSh_Style.Range("J" & line).Borders(xlLeft).Weight
            .Borders(xlLeft).Color = LadexSh_Style.Range("J" & line).Borders(xlLeft).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlRight).LineStyle <> xlNone Then
            .Borders(xlRight).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlRight).LineStyle
            .Borders(xlRight).Weight = LadexSh_Style.Range("J" & line).Borders(xlRight).Weight
            .Borders(xlRight).Color = LadexSh_Style.Range("J" & line).Borders(xlRight).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlTop).LineStyle <> xlNone Then
            .Borders(xlTop).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlTop).LineStyle
            .Borders(xlTop).Weight = LadexSh_Style.Range("J" & line).Borders(xlTop).Weight
            .Borders(xlTop).Color = LadexSh_Style.Range("J" & line).Borders(xlTop).Color
          End If

          If LadexSh_Style.Range("J" & line).Borders(xlBottom).LineStyle <> xlNone Then
            .Borders(xlBottom).LineStyle = LadexSh_Style.Range("J" & line).Borders(xlBottom).LineStyle
            .Borders(xlBottom).Weight = LadexSh_Style.Range("J" & line).Borders(xlBottom).Weight
            .Borders(xlBottom).Color = LadexSh_Style.Range("J" & line).Borders(xlBottom).Color
          End If
        End If


        '�w�i�F
        If LadexSh_Style.Range("H" & line) = "TRUE" Then
          .Interior.Color = LadexSh_Style.Range("J" & line).Interior.Color
        End If
      End With
    End If
  Next
  On Error GoTo catchError
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �X�^�C��������()
  Dim FSO As Object
  Dim setActivBook     As Workbook
  Dim FilePath As String, fileName As String
  
  Const funcName As String = "Ctl_Style.�X�^�C��������"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("�X�^�C�����p�m�F")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Call Library.�X�^�C���폜

  Set setActivBook = ActiveWorkbook
  Set setStyleBook = Workbooks.add
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  '�V�K�u�b�N���쐬���X�^�C�����C���|�[�g--------
  With setStyleBook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      FilePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs FilePath
  End With
  Call Library.showDebugForm("�V�K�u�b�N�쐬", FilePath, "debug")
  Call Library.showDebugForm("�X�^�C���̃}�[�W", , "debug")
  
  
  setActivBook.Activate
  ActiveWorkbook.Styles.Merge Workbook:=Workbooks(fileName)
  Set FSO = Nothing
  setStyleBook.Close
  
  Call Library.execDel(FilePath)
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function �W���X�^�C���̌����ڕύX()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Style.�W���X�^�C���̌����ڕύX"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  With ActiveWorkbook.Styles("Percent")
    .IncludeNumber = True
    .IncludeFont = True
    .IncludeAlignment = True
    .IncludeBorder = True
    .IncludePatterns = True
    .IncludeProtection = True
  
    '�t�H���g
    .Font.Name = "���C���I"
    .Font.Size = 9
    .Font.Color = ""
    .Font.Bold = ""
  
    '�z�u
    .HorizontalAlignment = ""
    .VerticalAlignment = ""
  
    '�r��
    .Borders(xlDiagonalDown).LineStyle = ""
    .Borders(xlDiagonalDown).Weight = ""
    .Borders(xlDiagonalDown).Color = ""
    .Borders(xlDiagonalUp).LineStyle = ""
    .Borders(xlDiagonalUp).Weight = ""
    .Borders(xlDiagonalUp).Color = ""
  
    .Borders(xlLeft).LineStyle = ""
    .Borders(xlLeft).Weight = ""
    .Borders(xlLeft).Color = ""
  
    .Borders(xlRight).LineStyle = ""
    .Borders(xlRight).Weight = ""
    .Borders(xlRight).Color = ""
  
    .Borders(xlTop).LineStyle = ""
    .Borders(xlTop).Weight = ""
    .Borders(xlTop).Color = ""
  
    .Borders(xlBottom).LineStyle = ""
    .Borders(xlBottom).Weight = ""
    .Borders(xlBottom).Color = ""
  
    '�w�i�F
    .Interior.Color = ""
   End With
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function �X�^�C���m�F()
  Dim styleCnt As Long
  Dim objSheet As Variant
  Dim objStyle As Variant
  Dim sheetName As String, styleName As String
  Dim slctRange As Range
  Dim RangeCnt As Long, RangeAllCnt As Long
  Dim chkNewSheetFlg As Boolean
  
  
  Const funcName As String = "Ctl_Style.�X�^�C���m�F"
  Const AbortCnt As Long = 10000
     
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Ctl_ProgressBar.showStart("�X�^�C�����p�m�F")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  styleCnt = 1
  chkNewSheetFlg = False
  Set useStyleVal = Nothing
  Set useStyleVal = CreateObject("Scripting.Dictionary")
  
  '���p�X�^�C���̎擾---------------------------
  Call Library.�X�^�C�����p�m�F
  
  For Each objStyle In ActiveWorkbook.Styles
    styleName = objStyle.Name
    Call Library.showDebugForm("�X�^�C����", styleName, "debug")
    
    If Library.chkArrayVal(useStyle(), styleName) = True Then
      Select Case styleName
        Case "Normal", "Percent", "Comma [0]", "Currency [0]", "Currency", "Comma"
        Case Else
          For Each objSheet In ActiveWorkbook.Sheets
            sheetName = objSheet.Name
            If Worksheets(sheetName).Visible = xlSheetVisible Then
              Worksheets(sheetName).Select
              Cells(Rows.count, Columns.count).Select
            
              chkNewSheetFlg = True
              RangeCnt = 1
              RangeAllCnt = Worksheets(sheetName).UsedRange.count
              
              For Each slctRange In Worksheets(sheetName).UsedRange
                If styleName = slctRange.style Then
                  If chkNewSheetFlg = True Then
                    slctRange.Select
                    chkNewSheetFlg = False
                  Else
                    Application.Union(Selection, slctRange).Select
                  End If
                End If
                'Call Library.showDebugForm("���p�X�^�C��", slctRange.style, "debug")
                Call Ctl_ProgressBar.showBar("�X�^�C���m�F", styleCnt, ActiveWorkbook.Styles.count, RangeCnt, RangeAllCnt, styleName)
                If RangeCnt >= AbortCnt Then
                  Exit For
                End If
                RangeCnt = RangeCnt + 1
              Next
              
              If Selection.Address <> Cells(Rows.count, Columns.count).Address Then
                If useStyleVal.Exists(styleName) Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>" & sheetName & "!" & Selection.Address
                Else
                  useStyleVal.add styleName, sheetName & "!" & Selection.Address
                End If
                If RangeCnt >= AbortCnt Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>Abort" & "!" & sheetName & "�V�[�g�ł̐ݒ萔���������݂��邽�߁A�����𒆒f"
                End If
              ElseIf Selection.Address = Cells(Rows.count, Columns.count).Address Then
              Else
                If useStyleVal.Exists(styleName) Then
                  useStyleVal(styleName) = useStyleVal(styleName) & "<|>Abort" & "!" & sheetName & "�V�[�g�ł̐ݒ萔���������݂��邽�߁A�����𒆒f"
                Else
                  useStyleVal.add styleName, "Abort" & "!" & sheetName & "�V�[�g�ł̐ݒ萔���������݂��邽�߁A�����𒆒f"
                End If
              End If
            End If
          Next
      End Select
      styleCnt = styleCnt + 1
    End If
  Next
  
  Call Ctl_ProgressBar.showEnd
  If useStyleVal.count > 0 Then
    Frm_Style.Show vbModeless
  Else
    Call Library.showNotice(10, "���p����Ă���X�^�C����������܂���ł���")
  End If
  
  Call Ctl_Book.A1�Z���I��

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
'    Call init.resetGlobalVal
  End If
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
