Attribute VB_Name = "Ctl_Book"
Option Explicit

'**************************************************************************************************
' * �u�b�N�Ǘ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �V�[�g�Ǘ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim addSheetName As String
  
  Const funcName As String = "Ctl_Book.�V�[�g�Ǘ�"
  
  '�����J�n--------------------------------------
  Application.Cursor = xlWait
  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  Frm_Sheet.Show vbModeless
  
  '�����I��--------------------------------------
  Call Library.endScript
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function R1C1�\�L()

  Const funcName As String = "Ctl_Book.R1C1�\�L"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
    Call Library.showDebugForm("R1C1�`���ɐݒ�", , "debug")
  Else
    Application.ReferenceStyle = xlA1
    Call Library.showDebugForm("A1�`���ɐݒ�", , "debug")
  End If
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
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
Function �W�����()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim SelectAddress, setZoomLevel, resetBgColor, setGgridLine
  
  Const funcName As String = "Ctl_Book.�W�����"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  PrgP_Max = 2
  PrgP_Cnt = PrgP_Cnt + 1
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  SelectAddress = Selection.Address
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  resetBgColor = Library.getRegistry("Main", "bgColor")
  setGgridLine = Library.getRegistry("Main", "gridLine")
  
  sheetCount = 1
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Worksheets(sheetName).Select
      
      '�W����ʂɐݒ�
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      ActiveWindow.View = xlNormalView
      
      '�\���{���̎w��
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      '�K�C�h���C���̕\��/��\��
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      If setGgridLine = "�\�����Ȃ�" Then
        ActiveWindow.DisplayGridlines = False
      ElseIf setGgridLine = "�\������" Then
        ActiveWindow.DisplayGridlines = True
      ElseIf setGgridLine = "�ύX���Ȃ�" Then
        'ActiveWindow.DisplayGridlines = setGgridLine
      End If
  
      '����͈͂̓_�����\��
      objSheet.DisplayAutomaticPageBreaks = False
      
      '�w�i�����Ȃ��ɂ���
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
      
      If resetBgColor = True Then
        With Application.FindFormat.Interior
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        With Application.ReplaceFormat.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
      End If
      
      'A1��I�����ꂽ��Ԃɂ���
      Application.GoTo Reference:=Range("A1"), Scroll:=True
      
      'RC�\�L����AQ�\�L�֕ύX
      If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
      End If
      
    End If
    Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select

  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd

  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function ���O��`�폜()
  Dim wb As Workbook, tmp As String
  Const funcName As String = "Ctl_Book.���O��`�폜"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call Library.�Z���̖��̐ݒ�폜
    

 '�����I��--------------------------------------
  If runFlg = False Then
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function �V�[�g�ꗗ�擾()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.�V�[�g�ꗗ�擾"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    If sheetNameLists = "" Then
      sheetNameLists = tempSheet.Name
    Else
      sheetNameLists = sheetNameLists & vbNewLine & tempSheet.Name
    End If
  Next
  
  With Frm_Info
    .Caption = "�V�[�g�ꗗ"
    .TextBox.Value = sheetNameLists
    .Show
  End With

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function ����͈͕\��()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.����͈͕\��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = False
  Next
  

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function ����͈͔�\��()
  Dim tempSheet As Object
  Dim sheetNameLists As String
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Book.����͈͔�\��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each tempSheet In Sheets
    tempSheet.DisplayAutomaticPageBreaks = True
  Next
  

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function �A���V�[�g�ǉ�()
  Dim sheetName As Variant
  
  Const funcName As String = "Ctl_Book.�A���V�[�g�ǉ�"

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
  '----------------------------------------------
  
  Set FrmVal = Nothing
  Set FrmVal = CreateObject("Scripting.Dictionary")
  With Frm_Info
    .Caption = "�A���V�[�g����"
    .TextBox.Value = ""
    .copySheet.Visible = True
    .Label1.Visible = True
    .Label2.Visible = True
    .Show
  End With
  
  Call Library.showDebugForm("�R�s�[��", FrmVal("copySheet"), "debug")
  
  For Each sheetName In Split(FrmVal("SheetList"), vbNewLine)
    Call Library.showDebugForm("�R�s�[��", sheetName, "debug")
    
    If Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") <> "��V�K�V�[�g��" Then
      Worksheets(FrmVal("copySheet")).copy After:=Worksheets(Worksheets.count)
      ActiveSheet.Name = CStr(sheetName)
    
    ElseIf Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") = "��V�K�V�[�g��" Then
      Worksheets.add(After:=Worksheets(Worksheets.count)).Name = CStr(sheetName)
    
    ElseIf Library.chkSheetExists(CStr(sheetName)) = True Then
      Call Library.showDebugForm("�R�s�[��쐬�ς�", sheetName, "debug")
    
    End If
    
    Application.GoTo Reference:=Range("A1"), Scroll:=True
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function A1�Z���I��()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim setZoomLevel  As Long
  
  Const funcName As String = "Ctl_Book.A1�Z���I��"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      
      ActiveWindow.Zoom = setZoomLevel
      Application.GoTo Reference:=Worksheets(sheetName).Range("A1"), Scroll:=True
    End If
    
    Call Ctl_ProgressBar.showBar("A1�Z���I��", 1, 2, sheetCount + 1, sheetMaxCount + 1, "�V�[�g�F" & sheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function A1�Z���I��_�ۑ�()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim setZoomLevel  As Long
  
  Const funcName As String = "Ctl_Book.A1�Z���I��_�ۑ�"

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
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      
      ActiveWindow.Zoom = setZoomLevel
      Application.GoTo Reference:=Worksheets(sheetName).Range("A1"), Scroll:=True
    End If
    
    Call Ctl_ProgressBar.showBar("A1�Z���I��", 1, 2, sheetCount + 1, sheetMaxCount + 1, "�V�[�g�F" & sheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  ActiveWorkbook.Save
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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

