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
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Book.���O��`�폜"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  
  For Each wb In Workbooks
    Workbooks(wb.Name).Activate
    Call Library.delVisibleNames
  Next wb
  
  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function

'==================================================================================================
Function �V�[�g���X�g�擾()
  Dim tempSheet As Object
  Dim infoVal As String
  Dim topPosition As Long, leftPosition As Long
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Book.�V�[�g���X�g�擾"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    If infoVal = "" Then
      infoVal = tempSheet.Name
    Else
      infoVal = infoVal & vbNewLine & tempSheet.Name
    End If
    
  Next

  topPosition = Library.getRegistry("UserForm", "InfoTop")
  leftPosition = Library.getRegistry("UserForm", "InfoLeft")
  
  Call Ctl_UsrForm.�\���ʒu(topPosition, leftPosition)
  With Frm_Info
    .StartUpPosition = 0
    .Top = topPosition
    .Left = leftPosition
    .TextBox.Value = infoVal
    .Show
  End With


  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function
