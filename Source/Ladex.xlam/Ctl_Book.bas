Attribute VB_Name = "Ctl_Book"
Option Explicit

'**************************************************************************************************
' * �u�b�N�Ǘ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function R1C1�\�L()

  Const funcName As String = "Ctl_Book.R1C1�\�L"
  
  '�����J�n--------------------------------------
  On Error Resume Next
  'Call init.setting
  '----------------------------------------------
  
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
  Else
    Application.ReferenceStyle = xlA1
  End If
  
  
  '�����I��--------------------------------------
  '----------------------------------------------
  Exit Function

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
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "function")
  End If
  PrgP_Max = 4
  PrgP_Cnt = 2
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  SelectAddress = Selection.Address
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  resetBgColor = Library.getRegistry("Main", "bgColor")
  setGgridLine = Library.getRegistry("Main", "gridLine")
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("SheetName", sheetName, "debug")
      
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
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------
  Exit Function

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
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
     Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
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
    Call Library.showDebugForm(funcName, , "start1")
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
    Call Library.showDebugForm(funcName, , "start1")
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
Function �A���V�[�g�ǉ�()
  Dim sheetName As Variant
  
  Const funcName As String = "Ctl_Book.�A���V�[�g�ǉ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
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
  
  Call Library.showDebugForm("copySheet", FrmVal("copySheet"), "debug")
  
  For Each sheetName In Split(FrmVal("SheetList"), vbNewLine)
    Call Library.showDebugForm("sheetName", sheetName, "debug")
    
    If Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") <> "��V�K�V�[�g��" Then
      Worksheets(FrmVal("copySheet")).copy After:=Worksheets(Worksheets.count)
      ActiveSheet.Name = CStr(sheetName)
    
    ElseIf Library.chkSheetExists(CStr(sheetName)) = False And sheetName <> "" And FrmVal("copySheet") = "��V�K�V�[�g��" Then
      Worksheets.add(After:=Worksheets(Worksheets.count)).Name = CStr(sheetName)
    End If
    
    Application.GoTo Reference:=Range("A1"), Scroll:=True
  Next
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function

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
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
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
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, sheetCount + 1, sheetMaxCount + 1, sheetName & "A1�Z���I��")
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

