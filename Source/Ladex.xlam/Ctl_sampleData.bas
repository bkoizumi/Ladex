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
      
      Case "�y���l�z�����Œ�"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame1.Caption = showType
      
      Case "�y���l�z�͈͎w��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame2.Caption = showType
      
      Case "�y���O�z��", "�y���O�z��", "�y���O�z�t���l�[��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        .MultiPage1.Pages.Item(5).Visible = False
        
        .Frame3.Caption = showType
        
      Case "�y���t�z��", "�y���t�z����", "�y���t�z����"
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

'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �p�^�[���I��()
  Dim line As Long, endLine As Long, count As Long
  Dim varDic
  Const funcName As String = "Library.�p�^�[���I��"
  
  '�����J�n--------------------------------------
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
  
  Call showFrm_sampleData("�p�^�[���I��")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  If sampleDataList Is Nothing Then
    End
  End If
  For Each varDic In sampleDataList
    Select Case sampleDataList.Item(varDic)
      Case "0.����(��)"
        Call ���O_��(maxCount)
      
      Case "1.����(��)"
        Call ���O_��(maxCount)
        
      Case "1.����(�t���l�[��)"
        Call ���O_�t���l�[��(maxCount)
      
      Case Else
    End Select
    ActiveCell.Offset(0, 1).Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
    Call Library.errorHandle
End Function

'==================================================================================================
Function ���l_�����Œ�(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���O��`�폜"
  
  '�����J�n--------------------------------------
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
  Call showFrm_sampleData("�y���l�z�����Œ�")
  
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "G/�W��"
    Cells(line + count, ActiveCell.Column) = Library.makeRandomDigits(BK_setVal("digits"))
  Next

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���l_�͈�()
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���l_�͈�"
  
  '�����J�n--------------------------------------
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
  
  Call showFrm_sampleData("�y���l�z�͈͎w��")
  line = Selection(1).Row
  maxCount = BK_setVal("maxCount")
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(BK_setVal("minVal"), BK_setVal("maxVal"))
  Next

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���O_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���O��`�폜"
  
  '�����J�n--------------------------------------
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
    Call showFrm_sampleData("�y���O�z��")
    maxCount = BK_setVal("maxCount")
  End If
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine)
  Next
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���O_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���O_��"
  
  '�����J�n--------------------------------------
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
    Call showFrm_sampleData("�y���O�z��")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("D" & getLine)
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���O_�t���l�[��(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���O_�t���l�[��"
  
  '�����J�n--------------------------------------
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
    Call showFrm_sampleData("�y���O�z�t���l�[��")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine) & "�@" & BK_sheetTestData.Range("D" & getLine)
  Next

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���t_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_SampleData.���O_�t���l�[��"
  
  '�����J�n--------------------------------------
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
  
  Call showFrm_sampleData("�y���t�z��")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row

  fstDate = BK_setVal("minVal")
  lstDate = BK_setVal("maxVal")
  
  For count = 0 To (maxCount - 1)
    Randomize
    Cells(line + count, ActiveCell.Column) = Format(Int((lstDate - fstDate + 1) * Rnd + fstDate), "yyyy/mm/dd")
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "yyyy/mm/dd"
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���t_����(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  Const funcName As String = "Ctl_SampleData.���t_����"
  
'  On Error GoTo catchError
  Call Library.startScript
  
  Call init.setting
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("�y���t�z����")
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
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function ����(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  Const funcName As String = "Ctl_SampleData.���t_����"
  
  '�����J�n--------------------------------------
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
  
  Call showFrm_sampleData("�y���t�z��")
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
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���̑�_����(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim makeStr As String
  Dim slctRange As Range
  Const funcName As String = "Ctl_SampleData.���̑�_���p����"
  
  '�����J�n--------------------------------------
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
  
  Call showFrm_sampleData("�y���̑��z����")
  
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
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

