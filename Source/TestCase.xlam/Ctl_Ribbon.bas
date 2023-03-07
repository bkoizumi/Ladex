Attribute VB_Name = "Ctl_Ribbon"

Public ribbonUI As IRibbonUI ' ���{��
Private rbButton_Visible As Boolean ' �{�^���̕\���^��\��
Private rbButton_Enabled As Boolean ' �{�^���̗L���^����

'�g�O���{�^��------------------------------------
Public PressT_B015 As Boolean


'**************************************************************************************************
' * ���{�����j���[�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�ǂݍ��ݎ�����------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  
  '���{���̕\�����X�V����
  ribbonUI.Invalidate
'  ribbonUI.ActivateTab "TestCaseTab"

End Function


'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '�����J�n--------------------------------------
'  runFlg = True
'  On Error GoTo catchError
'  Call init.setting(True)
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.startScript
'  PrgP_Max = 4
'  resetCellFlg = True
'  runFlg = True
'  Call Ctl_ProgressBar.showStart
  '----------------------------------------------

  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "CloseAddOn"
      Call Library.addonClose
    
    Case "OptionAddin����"
      Workbooks(ThisWorkbook.Name).IsAddin = False
    
    '�e�X�g�d�l��--------------------------------
    Case "�V�[�g�ǉ�"
      Call Ctl_TestCase.�V�[�g�ǉ�
      
    Case "�����ݒ�"
      Call Ctl_TestCase.�Đݒ�
      
      
    'selenium------------------------------------
    Case "Selenium���s"
      Call Ctl_Selenium.�J�n
      
      
      
      
      
    Case Else
      Call Library.showDebugForm("���{�����j���[�Ȃ�", control.ID, "Error")
      Call Library.showNotice(406, "���{�����j���[�Ȃ��F" & control.ID, True)
  End Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end")
  Call init.unsetting
  '----------------------------------------------
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
