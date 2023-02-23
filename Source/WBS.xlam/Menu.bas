Attribute VB_Name = "Menu"
'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting
  helpSheet.Visible = True
  helpSheet.Select
End Sub



'==================================================================================================
Sub optionKey()
Attribute optionKey.VB_ProcData.VB_Invoke_Func = "O\n14"
  Call M_�I�v�V������ʕ\��
End Sub
Sub centerKey()
  Call M_�Z���^�[
End Sub
Sub filterKey()
  Call M_�t�B���^�[
End Sub
Sub clearFilterKey()
  Call M_���ׂĕ\��
End Sub
Sub taskCheckKey()
Attribute taskCheckKey.VB_ProcData.VB_Invoke_Func = "C\n14"
  Call M_�^�X�N�`�F�b�N
End Sub
Sub makeGanttKey()
Attribute makeGanttKey.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call M_�K���g�`���[�g����
End Sub
Sub clearGanttKey()
Attribute clearGanttKey.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call M_�K���g�`���[�g�N���A
End Sub
Sub dispAllKey()
  Call M_���ׂĕ\��
End Sub
Sub taskControlKey()
'  Call M_
End Sub
Sub ScaleKey()
'  Call M_
End Sub








Sub M_�I�v�V������ʕ\��()
Attribute M_�I�v�V������ʕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  
  Call init.setting(True)
  Call Library.startScript
  
  Call WBS_Option.�I�v�V������ʕ\��
  
'  Call M_�J�����_�[����(True)
'  Call M_�K���g�`���[�g����
'  Call WBS_Option.�\����ݒ�
'
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
End Sub


Sub M_����ւ�()
  Call init.setting
  
  Call Library.startScript
  Call Check.���ڗ�`�F�b�N
  Call init.setting(True)
  
  Call Library.endScript
End Sub

Sub M_�J�����_�[����(Optional flg As Boolean = False)

  Call init.setting(True)
  Call Library.startScript
  
  If flg = False Then
    Call ProgressBar.showStart
  End If
  
  '�S�Ă̍s���\��
  Cells.EntireColumn.Hidden = False
  Cells.EntireRow.Hidden = False
  
  Call Ctl_Calendar.�J�����_�[����
  
  Call WBS_Option.�����̒S���ҍs���\��
  Call WBS_Option.�\����ݒ�
  
  If flg = False Then
    Call ProgressBar.showEnd
  End If
  Call Library.endScript
End Sub




'**************************************************************************************************
' * ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�s�n�C���C�g()
  Call Library.startScript
  Call WBS_Option.setLineColor
  Call Library.endScript
End Sub


'--------------------------------------
Sub M_�S�f�[�^�폜()
  If MsgBox("�f�[�^���폜���܂�", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  
  Call Library.startScript
  Call WBS_Option.clearAll
  Call Library.endScript
End Sub


Sub M_�S���()
Attribute M_�S���.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.ScreenUpdating = False
  ActiveWindow.DisplayHeadings = False
  Application.DisplayFullScreen = True
  
  With DispFullScreenForm
    .StartUpPosition = 0
    .Top = Application.Top + 300
    .Left = Application.Left + 30
  End With
  Application.ScreenUpdating = True
  DispFullScreenForm.Show vbModeless
End Sub

Sub M_�^�X�N����()
Attribute M_�^�X�N����.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

Sub M_�X�P�[��()
Attribute M_�X�P�[��.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @Link   https://akashi-keirin.hatenablog.com/entry/2019/02/23/170807
'**************************************************************************************************
Sub M_�^�X�N�`�F�b�N()
Attribute M_�^�X�N�`�F�b�N.VB_ProcData.VB_Invoke_Func = "C\n14"
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Dim targetBookName As String
  Dim targetSheet As Object
  
  Const funcName As String = "Menu.M_�^�X�N�`�F�b�N"

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
  
  targetBookName = ActiveWorkbook.Name
  Set targetSheet = ActiveWorkbook.Worksheets("WBS")
  
  PrgP_Cnt = PrgP_Cnt + 1
  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")
  
  
  Call Ctl_Task.�^�X�N�`�F�b�N
  
  
  '�_����쐬VBA�̌Ăяo��
  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "!sheet5.CommandButton7_Click"
  
  Set targetSheet = Nothing
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
Sub M_�t�B���^�[()
Attribute M_�t�B���^�[.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  
  With FilterForm
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
  End With
  
  FilterForm.Show
End Sub


'==================================================================================================
Sub M_���ׂĕ\��()
Attribute M_���ׂĕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  Call Library.startScript
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call WBS_Option.�����̒S���ҍs���\��
  Call Library.endScript
End Sub


'==================================================================================================
Sub M_�i���R�s�[()
  Call Task.�i���R�s�[
End Sub

'==================================================================================================
Sub M_�^�X�N�ړ�_��()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�ړ�_��
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�ړ�_��()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�ړ�_��
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�ړ�_��()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�ړ�_��
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�ړ�_�E()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�ړ�_�E
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�ǉ�()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�ǉ�
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�폜()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.�^�X�N�폜
  Call Library.endScript
End Sub












'==================================================================================================
Sub M_�i�����ݒ�(progress As Long)
  Call Task.�i�����ݒ�(progress)
End Sub

'�^�X�N�̃����N�ݒ�/����---------------------------------------------------------------------------
Sub M_�^�X�N�̃����N�ݒ�()
  Call Ctl_Task.�^�X�N�̃����N�ݒ�
End Sub

'==================================================================================================
Sub M_�^�X�N�̃����N����()
  Call Ctl_Task.�^�X�N�̃����N����
End Sub

Sub M_�^�X�N�̑}��()
  Call Library.startScript
  Call init.setting
  
  Call Task.�^�X�N�̑}��
  
  Call Library.endScript
End Sub

Sub M_�^�X�N�̍폜()
  Call Library.startScript
  Call init.setting
  
  Call Task.�^�X�N�̍폜
  
  Call Library.endScript
End Sub

'�\�����[�h----------------------------------------------------------------------------------------
Sub M_�^�X�N�\��_�W��()
  Call Library.startScript
  
  Call init.setting
  If setVal("debugMode") <> "develop" Then
    mainSheet.Visible = True
    TeamsPlannerSheet.Visible = xlSheetVeryHidden
  End If
  
  Call init.setting(True)
  Call WBS_Option.�^�X�N�\��_�W��
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  Call Library.endScript

End Sub

Sub M_�^�X�N�\��_�^�X�N()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTask
  Call WBS_Option.setLineColor
  
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_�^�X�N�\��_�`�[���v�����i�[()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.�^�X�N�\��_�`�[���v�����i�[
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
  Call Library.endScript
End Sub


'==================================================================================================
Sub M_�^�X�N�ɃX�N���[��()
  Const funcName As String = "Menu.M_�^�X�N�ɃX�N���[��"

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
  
  
  Call Task.�^�X�N�ɃX�N���[��
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub




'==================================================================================================
Sub M_�^�C�����C���ɒǉ�()
  Const funcName As String = "Menu.M_�^�C�����C���ɒǉ�"

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
  
  Call Ctl_Chart.�^�C�����C���ɒǉ�(ActiveCell.Row)
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'**************************************************************************************************
' * �K���g�`���[�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�N���A--------------------------------------------------------------------------------------------
Sub M_�K���g�`���[�g�N���A()
Attribute M_�K���g�`���[�g�N���A.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call Library.startScript
  Call Ctl_Chart.�K���g�`���[�g�폜
  Call Library.endScript
End Sub

'�����̂�------------------------------------------------------------------------------------------
Sub M_�K���g�`���[�g�����̂�()
Attribute M_�K���g�`���[�g�����̂�.VB_ProcData.VB_Invoke_Func = "A\n14"
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("�K���g�`���[�g����", "�����J�n")
  
  Call Ctl_Chart.�K���g�`���[�g����
  
  Call Library.showDebugForm("�K���g�`���[�g����", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'==================================================================================================
'����
Sub M_�K���g�`���[�g����()
Attribute M_�K���g�`���[�g����.VB_ProcData.VB_Invoke_Func = "t\n14"
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim targetBookName As String
  Dim SelectionSheet As String
  Dim objShp As Shape
  
  Const funcName As String = "Menu.M_�K���g�`���[�g����"

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
  
  targetBookName = ActiveWorkbook.Name
  
  '�_����쐬VBA�̌Ăяo��
  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet5.CommandButton1_Click"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.Title_Format_Check"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.CALC_MANUAL"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Chart_Module.Make_Chart"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.BUTTON_CLEAR"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")
  
  


  
  '�^�C�����C���ɒǉ�----------------------------
  Call Library.startScript
  Rows("6:6").RowHeight = 40
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "TimeLine_*" Then
      ActiveSheet.Shapes(objShp.Name).Delete
    End If
  Next
  
  
  For line = 2 To Sh_PARAM.Cells(Rows.count, 38).End(xlUp).Row
    If Sh_PARAM.Range("AL" & line).Text <> "" Then
      Call Ctl_Chart.�^�C�����C���ɒǉ�(CLng(Sh_PARAM.Range("AL" & line).Text), True)
    End If
  Next

  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'==================================================================================================
'�Z���^�[
Sub M_�Z���^�[()
Attribute M_�Z���^�[.VB_ProcData.VB_Invoke_Func = " \n14"

  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("�Z���^�[�ֈړ�", "�����J�n")
  
  Call Ctl_Chart.�Z���^�[
  
  Call Library.showDebugForm("�Z���^�[�ֈړ�", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excel�t�@�C��-------------------------------------------------------------------------------------
Sub M_Excel�C���|�[�g()
  
  Call Library.startScript
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "�����J�n")
  
  Call init.setting
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).Row
  
  If endLine > 6 Then
    If MsgBox("�f�[�^���폜���܂�", vbYesNo + vbExclamation) = vbYes Then
      Call WBS_Option.clearAll
    Else
      Call WBS_Option.clearCalendar
    End If
  Else
    Call WBS_Option.clearCalendar
  End If
  Call ProgressBar.showStart
  
  
  Call Import.�t�@�C���C���|�[�g
  Call Calendar.�����ݒ�
  Call Import.�J�����_�[�p�����擾
  
  If setVal("lineColorFlg") = "True" Then
    setVal("lineColorFlg") = False
    Call WBS_Option.setLineColor
  Else
  End If
  
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
  Call WBS_Option.saveAndRefresh
  
  Application.Goto Reference:=Range("A6"), Scroll:=True


  Err.Clear
  Call Library.showNotice(200, "�C���|�[�g")
End Sub


