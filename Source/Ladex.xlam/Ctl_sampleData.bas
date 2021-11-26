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
      
      Case "�y���l�z�����Œ�"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame1.Caption = showType
      
      Case "�y���l�z�͈͎w��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame2.Caption = showType
      
      Case "�y���O�z��", "�y���O�z��", "�y���O�z�t���l�[��"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(4).Visible = False
        
        .Frame3.Caption = showType
        
      Case "�y���t�z��", "�y���t�z����", "�y���t�z����"
        .MultiPage1.Pages.Item(0).Visible = False
        .MultiPage1.Pages.Item(1).Visible = False
        .MultiPage1.Pages.Item(2).Visible = False
        .MultiPage1.Pages.Item(3).Visible = False
        
        .minVal4 = #4/1/2021#
        .maxVal4 = #3/31/2022#
        
        .Frame4.Caption = showType
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
  
  
'  On Error GoTo catchError
  Call init.setting
  
'  Sheets("Sheet1").Select
'  Sheets("Sheet1").Columns("A:Z").Clear
'  Sheets("Sheet1").Range("A1").Select
  
  Call showFrm_sampleData("�p�^�[���I��")
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  If sampleDataList Is Nothing Then
    End
  End If
  For Each varDic In sampleDataList
    Debug.Print sampleDataList.Item(varDic)
    
    Select Case sampleDataList.Item(varDic)
      Case "0.����(��)"
        Call ���O_��(maxCount)
      
      Case "1.����(��)"
        Call ���O_��(maxCount)
        
      Case "1.����(�t���l�[��)"
        Call ���O_�t���l�[��(maxCount)
      
      Case Else
    End Select
    '���̃Z���Ɉړ�
    ActiveCell.Offset(0, 1).Select
  Next
  
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
    Call Library.errorHandle
End Function


'==================================================================================================
Function ���l_�����Œ�(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call init.setting
  
  Call showFrm_sampleData("�y���l�z�����Œ�")
  
  maxCount = BK_setVal("maxCount")
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column).NumberFormatLocal = "G/�W��"
    Cells(line + count, ActiveCell.Column) = Library.makeRandomDigits(BK_setVal("digits"))
  Next


  Exit Function
'�G���[������====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���l_�͈�()
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  
'  Sheets("Sheet1").Columns("B:B").Clear
'  Sheets("Sheet1").Range("B1").Select
  
  
  Call showFrm_sampleData("�y���l�z�͈͎w��")
  line = Selection(1).Row
  maxCount = BK_setVal("maxCount")
  
  For count = 0 To (maxCount - 1)
    Cells(line + count, ActiveCell.Column) = Library.makeRandomNo(BK_setVal("minVal"), BK_setVal("maxVal"))
  Next

  Call Library.endScript
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function ���O_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  
  Call Library.startScript
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("C:C").Clear
'  Sheets("Sheet1").Range("C1").Select
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("�y���O�z��")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine)
  Next

  Call Library.endScript
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function ���O_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("D:D").Clear
'  Sheets("Sheet1").Range("D1").Select
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("�y���O�z��")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("D" & getLine)
  Next

  Call Library.endScript
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function

'==================================================================================================
Function ���O_�t���l�[��(Optional maxCount As Long)
  Dim line As Long, endLine As Long

'  On Error GoTo catchError
  Call Library.startScript
  
  Call init.setting
  endLine = BK_sheetTestData.Cells(Rows.count, 1).End(xlUp).Row
  
'  Sheets("Sheet1").Columns("E:E").Clear
'  Sheets("Sheet1").Range("E1").Select
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("�y���O�z�t���l�[��")
    maxCount = BK_setVal("maxCount")
  End If
  
  line = Selection(1).Row
  
  For count = 0 To (maxCount - 1)
    getLine = Library.makeRandomNo(2, endLine)
    Cells(line + count, ActiveCell.Column) = BK_sheetTestData.Range("A" & getLine) & "�@" & BK_sheetTestData.Range("D" & getLine)
  Next
  Call Library.endScript

  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function ���t_��(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  
  Call Library.startScript
  
'  On Error GoTo catchError
  
  Call init.setting
  
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
  
  Call Library.endScript
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function ���t_����(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call Library.startScript
  
  Call init.setting
  
  If maxCount <= 1 Then
    Call showFrm_sampleData("�y���t�z����")
  End If
  
  maxCount = BK_setVal("maxCount")
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
End Function



'==================================================================================================
Function ����(Optional maxCount As Long)
  Dim line As Long, endLine As Long
  Dim val As Double
  
'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  
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
    
'    Cells(line + count, ActiveCell.Column + 1) = DateAdd("s", Library.makeRandomNo(0, 600), val)
'    Cells(line + count, ActiveCell.Column + 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
  Next
  
  Call Library.endScript
  
  Exit Function
'�G���[������====================================
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function























