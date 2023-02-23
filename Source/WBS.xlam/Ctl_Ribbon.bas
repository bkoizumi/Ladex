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
  ribbonUI.ActivateTab "WBSTab"

End Function


'==================================================================================================
'�g�O���{�^���Ƀ`�F�b�N��ݒ肷��
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
  Select Case control.ID
    Case "T_B015"
      '�^�C�����C���ɒǉ�
      If Range(setVal("cell_Info") & ActiveCell.Row) Like "" Then
        returnedVal = True
      Else
        returnedVal = False
      End If
      
      
      
    Case Else
  End Select
End Sub


'==================================================================================================
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 2)
End Sub

'==================================================================================================
Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.ID, 3)
  Application.run setRibbonVal

End Sub


'==================================================================================================
'Supertip�̓��I�\��
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 5)
End Sub

'==================================================================================================
Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 6)
End Sub

'==================================================================================================
Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.ID, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select


End Sub

'==================================================================================================
'Ribbon�V�[�g������e���擾
Function getRibbonMenu(menuId As String, offsetVal As Long)

  Dim getString As String
  Dim FoundCell As Range
  Dim ribSheet As Worksheet
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Set ribSheet = ThisWorkbook.Worksheets("Ribbon")

  endLine = ribSheet.Cells(Rows.count, 1).End(xlUp).Row

  getRibbonMenu = Application.VLookup(menuId, ribSheet.Range("A2:F" & endLine), offsetVal, False)
  Call Library.endScript


  Exit Function
'�G���[������=====================================================================================
catchError:
  getRibbonMenu = "�G���["

End Function


'**************************************************************************************************
' * ���{��������s
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '�����J�n--------------------------------------
  runFlg = True
'  On Error GoTo catchError
  Call init.setting(True)
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  PrgP_Max = 4
  resetCellFlg = True
  '----------------------------------------------
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "CloseAddOn"
      Call Library.addonClose
    
    Case "OptionAddin����"
      Workbooks(ThisWorkbook.Name).IsAddin = False
    
    Case "OptionAddin��"
      Workbooks(ThisWorkbook.Name).IsAddin = True
      ThisWorkbook.Save
    
    
    
    '�i�����ݒ�----------------------------------
    Case "progress_0", "progress_25", "progress_50", "progress_75", "progress_100"
      Call Ctl_Task.�i�����ݒ�(Replace(control.ID, "progress_", ""))
    
    
    '�^�X�N�ړ�----------------------------------
    Case "taskOutdent"
      Call Library.testCalAddDay
      
      Call Ctl_Task.�^�X�N�ړ�_��
      
    Case "taskIndent"
      Call Ctl_Task.�^�X�N�ړ�_�E
      
    Case "taskLink"
      Call Ctl_Task.�^�X�N�̃����N�ݒ�
    Case "taskLink"
      Call Ctl_Task.�^�X�N�̃����N����
      
    Case "chkTaskList"
      Call Ctl_Task.�^�X�N�`�F�b�N
      
      '�_����쐬VBA�̌Ăяo��
      Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!sheet5.CommandButton7_Click"


    Case "Option"
      Call Ctl_Option.�I�v�V������ʕ\��
      
    Case "scrollTask"
      Call Ctl_Task.�^�X�N�ɃX�N���[��
    
    Case "addTimeLine"
      Call Ctl_Chart.�^�C�����C���ɒǉ�(ActiveCell.Row)
        
    Case "makeChart"
         Menu.M_�K���g�`���[�g����
         
    Case "makeCalendar"
      Call Ctl_Calendar.�J�����_�[����
      'Call Ctl_Task.�^�X�N�`�F�b�N

    Case "copyProgress"
      Call Ctl_Task.�i���R�s�[
      
      
    Case Else
      Call Library.showDebugForm("���{�����j���[�Ȃ�", control.ID, "Error")
      Call Library.showNotice(406, "���{�����j���[�Ȃ��F" & control.ID, True)
  End Select
  
  '�����I��--------------------------------------
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


