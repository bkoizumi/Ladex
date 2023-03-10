Attribute VB_Name = "Ctl_sampleData"
Option Explicit

Dim newBook As Workbook
Dim count As Long, getLine As Long
Dim fstDate As Date, lstDate As Date

Public maxCount  As Long

'**************************************************************************************************
' * �f�[�^����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showFrm_sampleData(showType As String)
  With Frm_smplData
    .Caption = showType
    
    '�e�y�[�W�A�p�[�c�̗L��/�����؂�ւ�
    Select Case showType
      Case "�p�^�[���I��"
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
      
      Case "�y���l�z�����Œ�"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
        
        .Frame1.Caption = showType
      
      Case "�y���l�z�͈͎w��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
        
        .Frame2.Caption = showType
      
      Case "�y���O�z��", "�y���O�z��", "�y���O�z�t���l�[��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
        
        .Frame3.Caption = showType
        
      Case "�y���t�z��", "�y���t�z����", "�y���t�z����"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
        
        .minVal4 = #4/1/2021#
        .maxVal4 = #3/31/2022#
        
        .Frame4.Caption = showType
        
        
      Case "�x�����X�g����"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .minVal4 = #4/1/2021#
        .maxVal4 = #3/31/2022#
        
        .Frame4.Caption = showType
        
        
      Case "�y���̑��z����"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(6).Visible = False
        
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
    If Selection.CountLarge > 1 Then
      .maxCount0 = Selection.Rows.count
      .maxCount1 = Selection.Rows.count
      .maxCount2 = Selection.CountLarge
      .maxCount3 = Selection.Rows.count
      .maxCount4 = Selection.Rows.count
    End If

    .Show
  End With
  
  Exit Function

'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_�p�^�[���I��()
  Dim line As Long, endLine As Long, count As Long, getLine As Long, getLine2 As Long
  Dim varDic
  Dim actRange As Range
  Dim strAddress As String
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_�p�^�[���I��"
  
  '�����J�n--------------------------------------
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
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  
  Call showFrm_sampleData("�p�^�[���I��")
  If sampleDataList Is Nothing Then
    End
  End If
  maxCount = dicVal("maxCount")

  Call Library.delSheetData(LadexSh_InputData)
  LadexSh_InputData.Cells.NumberFormatLocal = "@"
  
  line = 1
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row)
    getLine2 = Library.makeRandomNo(2, 5)
    
    '����(��)
    LadexSh_InputData.Range("A" & line + count) = LadexSh_TestData.Range("A" & getLine)
    LadexSh_InputData.Range("D" & line + count) = LadexSh_TestData.Range("B" & getLine)
    
    '����(��)
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 4).End(xlUp).Row)
    LadexSh_InputData.Range("B" & line + count) = LadexSh_TestData.Range("D" & getLine)
    LadexSh_InputData.Range("E" & line + count) = LadexSh_TestData.Range("E" & getLine)
    
    LadexSh_InputData.Range("C" & line + count) = LadexSh_InputData.Range("A" & line + count) & "�@" & LadexSh_InputData.Range("B" & line + count)
    LadexSh_InputData.Range("F" & line + count) = LadexSh_InputData.Range("D" & line + count) & "�@" & LadexSh_InputData.Range("E" & line + count)
    
    '����
    LadexSh_InputData.Range("G" & line + count) = LadexSh_TestData.Range("F" & getLine)
    
    '���t�^
    LadexSh_InputData.Range("H" & line + count) = LadexSh_TestData.Range("H" & getLine2)
    
    '���N����
    LadexSh_InputData.Range("I" & line + count) = Format(Int((Date - #1/1/1950# + 1) * Rnd + #1/1/1950#), "yyyy/mm/dd")
    
    '�N��
    LadexSh_InputData.Range("J" & line + count) = Application.Evaluate("DATEDIF(""" & LadexSh_InputData.Range("I" & line + count) & """, TODAY(), ""Y"")")
    
    '�d�b�ԍ�
    LadexSh_InputData.Range("K" & line + count) = LadexSh_TestData.Range("Z" & getLine) & "-" & LadexSh_TestData.Range("AA" & getLine) & "-1234"
    
    '���[���A�h���X
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 10).End(xlUp).Row)
    LadexSh_InputData.Range("L" & line + count) = "Sample" & LadexSh_TestData.Range("J" & getLine)
    
    '�s���{��
    getLine = Library.makeRandomNo(2, LadexSh_TestData.Cells(Rows.count, 15).End(xlUp).Row)
    LadexSh_InputData.Range("M" & line + count) = LadexSh_TestData.Range("O" & getLine)
    LadexSh_InputData.Range("N" & line + count) = LadexSh_TestData.Range("P" & getLine)
    LadexSh_InputData.Range("O" & line + count) = LadexSh_TestData.Range("Q" & getLine)
    LadexSh_InputData.Range("P" & line + count) = LadexSh_TestData.Range("R" & getLine)
    LadexSh_InputData.Range("Q" & line + count) = LadexSh_TestData.Range("S" & getLine)
   
    If InStr(LadexSh_TestData.Range("U" & getLine2), "��") > 0 Then
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & LadexSh_TestData.Range("T" & getLine) & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbWide)
    Else
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & StrConv(Replace(LadexSh_TestData.Range("T" & getLine), "����", "-"), vbNarrow)
      LadexSh_InputData.Range("R" & line + count) = LadexSh_InputData.Range("R" & line + count) & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbNarrow)
    End If
    
    LadexSh_InputData.Range("S" & line + count) = LadexSh_TestData.Range("V" & getLine)
    LadexSh_InputData.Range("T" & line + count) = LadexSh_TestData.Range("W" & getLine)
    LadexSh_InputData.Range("U" & line + count) = LadexSh_TestData.Range("X" & getLine)
    
    strAddress = LadexSh_InputData.Range("R" & line + count)
    strAddress = StrConv(Replace(strAddress, "����", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "��", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "�Ԓn", ""), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "��", "-"), vbNarrow)
    strAddress = StrConv(Replace(strAddress, "��", ""), vbNarrow)
    
    
    LadexSh_InputData.Range("V" & line + count) = strAddress
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  Set actRange = Selection(1)
  actRange.Select
  
  For Each varDic In sampleDataList
    Call Library.showDebugForm("varDic", varDic, "debug")
    Select Case varDic
      Case "����(��)"
        LadexSh_InputData.Range("A1:A" & maxCount).copy Selection
      Case "����(��)"
        LadexSh_InputData.Range("B1:B" & maxCount).copy Selection

      Case "����(�t���l�[��)"
        LadexSh_InputData.Range("C1:C" & maxCount).copy Selection

      Case "[�J�i]����(��)"
        LadexSh_InputData.Range("D1:D" & maxCount).copy Selection
      Case "[�J�i]����(��)"
        LadexSh_InputData.Range("E1:E" & maxCount).copy Selection
      Case "[�J�i]����(�t���l�[��)"
        LadexSh_InputData.Range("F1:F" & maxCount).copy Selection
      Case "����"
        LadexSh_InputData.Range("G1:G" & maxCount).copy Selection
      Case "���t�^"
        LadexSh_InputData.Range("H1:H" & maxCount).copy Selection
      Case "���N����"
        LadexSh_InputData.Range("I1:I" & maxCount).copy Selection
      Case "�N��"
        LadexSh_InputData.Range("J1:J" & maxCount).copy Selection
      Case "�d�b�ԍ�"
        LadexSh_InputData.Range("K1:K" & maxCount).copy Selection
      Case "���[���A�h���X"
        LadexSh_InputData.Range("L1:L" & maxCount).copy Selection
      Case "�s���{���R�[�h"
        LadexSh_InputData.Range("M1:M" & maxCount).copy Selection
      Case "�X�֔ԍ�"
        LadexSh_InputData.Range("N1:N" & maxCount).copy Selection
      Case "�s���{��"
        LadexSh_InputData.Range("O1:O" & maxCount).copy Selection
      Case "�s��S����"
        LadexSh_InputData.Range("P1:P" & maxCount).copy Selection
      Case "����"
        LadexSh_InputData.Range("Q1:Q" & maxCount).copy Selection
      Case "���ځE�����E�Ԓn"
        LadexSh_InputData.Range("R1:R" & maxCount).copy Selection
      Case "[�J�i]�s���{��"
        LadexSh_InputData.Range("S1:S" & maxCount).copy Selection
      Case "[�J�i]�s��S����"
        LadexSh_InputData.Range("T1:T" & maxCount).copy Selection
      Case "[�J�i]����"
        LadexSh_InputData.Range("U1:U" & maxCount).copy Selection
      Case "[�J�i]���ځE�����E�Ԓn"
        LadexSh_InputData.Range("V1:V" & maxCount).copy Selection
      
      Case Else
    End Select
    ActiveCell.Offset(0, 1).Select
    DoEvents
  Next
  actRange.Select
  Call Library.delSheetData(LadexSh_InputData)
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
    Call Library.errorHandle
End Function

'==================================================================================================
Function ���l_�����Œ�()
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.���l_�����Œ�"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  Call showFrm_sampleData("�y���l�z�����Œ�")
  
  If IsMissing(maxCount) Then
    maxCount = dicVal("maxCount")
  End If
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "###"
    Cells(line + count, ActiveCell.Column) = dicVal("addFirst") & Library.makeRandomDigits(dicVal("digits")) & dicVal("addEnd")
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���l_�͈͎w��()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.���l_�͈͎w��"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("�y���l�z�͈͎w��")
  
  If Selection.CountLarge > 1 Then
    For Each slctCells In Selection
      slctCells.NumberFormatLocal = "###"
      slctCells.Value = Library.makeRandomNo(dicVal("minVal"), dicVal("maxVal"))
      DoEvents
    Next
  Else
    line = Selection(1).Row
  
    If maxCount = 0 Then
      maxCount = dicVal("maxCount")
    End If
  
    For count = 0 To (maxCount - 1)
      Cells(line + count, ActiveCell.Column).NumberFormatLocal = "###"
      Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(dicVal("minVal"), dicVal("maxVal"))
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
    Next
  End If
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_��()
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_��"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
  If IsMissing(maxCount) Then
    Call showFrm_sampleData("�y���O�z��")
    maxCount = dicVal("maxCount")
  End If
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("A" & getLine)
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_��()
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_��"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  If IsMissing(maxCount) Then
'    Call showFrm_sampleData("�y���O�z��")
'    maxCount = dicVal("maxCount")
'  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("D" & getLine)
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_����(Optional kanaFlg As Boolean = False)
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_����"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  endLine = LadexSh_TestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  If IsMissing(maxCount) Then
'    Call showFrm_sampleData("�y���O�z�t���l�[��")
'    maxCount = dicVal("maxCount")
'  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("A" & getLine) & "�@" & LadexSh_TestData.Range("D" & getLine)
    
    If kanaFlg = True Then
      Cells(line + count, ActiveCell.Column + 1) = LadexSh_TestData.Range("B" & getLine) & "�@" & LadexSh_TestData.Range("E" & getLine)
    End If
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_���t()
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_���t"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  Call showFrm_sampleData("�y���t�z��")
'  If IsMissing(maxCount) Then
'    maxCount = dicVal("maxCount")
'  End If
  line = Selection(1).Row

  fstDate = dicVal("minVal")
  lstDate = dicVal("maxVal")
  
  For count = 0 To (maxCount - 1)
    Randomize
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd"
    Cells(line + count, ActiveCell.Column) = Format(Int((lstDate - fstDate + 1) * Rnd + fstDate), "yyyy/mm/dd")
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_����()
  Dim line As Long, endLine As Long
  Dim maxCount As Long
  Dim val As Double
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_����"
  
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
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
'  If IsMissing(maxCount) Then
'    Call showFrm_sampleData("�y���t�z����")
'    maxCount = dicVal("maxCount")
'  End If
  
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("00:00:00") * 100000, TimeValue("23:59:59") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = Format(val, "hh:nn:ss")
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "hh:mm:ss"
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_����()
  Dim line As Long, endLine As Long
  Dim val As Double
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_����"
  
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
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  maxCount = Selection.Rows.count
  Call showFrm_sampleData("�y���t�z��")
'  If IsMissing(maxCount) Then
'    maxCount = dicVal("maxCount")
'  End If
  line = Selection(1).Row

  fstDate = dicVal("minVal")
  lstDate = dicVal("maxVal")
  
  line = Selection(1).Row
  For count = 0 To maxCount - 1
    Randomize
    val = WorksheetFunction.RandBetween(TimeValue("09:00:00") * 100000, TimeValue("18:00:00") * 100000) / 100000
    val = Int((lstDate - fstDate + 1) * Rnd + fstDate) + val

    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = val
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_����()
  Dim line As Long, endLine As Long
  Dim makeStr As String, slctRange As Range
  Dim maxCount As Long
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_����"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call showFrm_sampleData("�y���̑��z����")
  If IsMissing(maxCount) Then
    maxCount = dicVal("maxCount")
  End If
  
  makeStr = ""
  If dicVal("strType01") Then makeStr = makeStr & SymbolCharacters
  If dicVal("strType02") Then makeStr = makeStr & HalfWidthCharacters
  If dicVal("strType03") Then makeStr = makeStr & StrConv(HalfWidthCharacters, vbLowerCase)
  If dicVal("strType04") Then makeStr = makeStr & HalfWidthDigit
  If dicVal("strType05") Then makeStr = makeStr & JapaneseCharacters
  If dicVal("strType06") Then makeStr = makeStr & StrConv(JapaneseCharacters, vbKatakana)
  If dicVal("strType07") Then makeStr = makeStr & MachineDependentCharacters

  For Each slctRange In Selection
    slctRange.Value = Library.makeRandomString(maxCount, makeStr)
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_���[���A�h���X()
  Dim line As Long, endLine As Long
  Dim makeStr As String
  Dim maxCount As Long
  Const funcName As String = "Ctl_SampleData.�f�[�^����_���[���A�h���X"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 10).End(xlUp).Row
  If IsMissing(maxCount) Then
    maxCount = dicVal("maxCount")
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
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_�Z��(maxCount As Long, ParamArray addressFlgs())
  Dim line As Long, endLine As Long
  Dim getLine As Long, getLine2 As Long
  Dim strAddress As String
  
  Const funcName As String = "Ctl_SampleData.�f�[�^����_�Z��"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 16).End(xlUp).Row
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    getLine2 = Library.makeRandomNo(2, 5)
    
    If InStr(LadexSh_TestData.Range("U" & getLine2), "��") > 0 Then
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
      strAddress = strAddress & StrConv(Replace(LadexSh_TestData.Range("T" & getLine), "����", "-"), vbUpperCase)
      strAddress = strAddress & StrConv(Replace(LadexSh_TestData.Range("U" & getLine2), "%", Library.makeRandomNo(1, 5)), vbNarrow)
    End If
    
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "@"
    Cells(line + count, ActiveCell.Column) = strAddress
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�[�^����_�d�b�ԍ�(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.�f�[�^����_�d�b�ԍ�"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  endLine = LadexSh_TestData.Cells(Rows.count, 15).End(xlUp).Row
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = LadexSh_TestData.Range("Y" & getLine) & "-" & LadexSh_TestData.Range("Z" & getLine) & "-1234"
  
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, count, maxCount, "�f�[�^����")
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �f�[�^����_�x�����X�g()
  Dim line As Long, endLine As Long
  Dim targetDay As Date, startDay As Date, endDay As Date
  Dim targetRange As Range
  Dim HollydayName As String
  Const funcName As String = "Ctl_SampleData.�f�[�^����_�x�����X�g"
  
  '�����J�n--------------------------------------
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  line = 0
  Set targetRange = ActiveCell
  
  Call showFrm_sampleData("�x�����X�g����")

  startDay = dicVal("minVal")
  endDay = dicVal("maxVal")
  
  
  For targetDay = #4/1/2022# To #3/31/2023#
    If Ctl_Hollyday.GetHollyday(targetDay, HollydayName) = True Then
      targetRange.Offset(line).Select
      targetRange.Offset(line) = targetDay
      line = line + 1
    End If
  Next
  targetRange.Select
  Set targetRange = Nothing
  
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
