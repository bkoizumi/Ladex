Attribute VB_Name = "Ctl_Book"
Option Explicit

'**************************************************************************************************
' * �u�b�N�Ǘ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function ���O��`�폜()
  Dim wb As Workbook, tmp As String
  Const funcName As String = "Ctl_Book.���O��`�폜"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
'  For Each wb In Workbooks
'    Workbooks(wb.Name).Activate
'    Call Library.delVisibleNames
'  Next wb
  
  Call Library.delVisibleNames
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �V�[�g���X�g�擾()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.�V�[�g���X�g�擾"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    If sheetNameLists = "" Then
      sheetNameLists = tempSheet.Name
    Else
      sheetNameLists = sheetNameLists & vbNewLine & tempSheet.Name
    End If
  Next
  
  With Frm_Info
    .TextBox.Value = sheetNameLists
    .Show
  End With

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ����͈͂̓_�����\��()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.����͈͂̓_�����\��"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = False
  Next
  

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ����͈͂̓_����\��()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.����͈͂̓_����\��"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = True
  Next
  

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
