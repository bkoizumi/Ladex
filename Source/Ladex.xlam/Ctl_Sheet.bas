Attribute VB_Name = "Ctl_Sheet"
Option Explicit

'**************************************************************************************************
' * R1C1�\�L
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function R1C1�\�L()

  On Error Resume Next
  
  Call init.setting
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
  Else
    Application.ReferenceStyle = xlA1
  End If
  
End Function



'==================================================================================================
Function A1�Z���I��()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  
  Const funcName As String = "Ctl_Sheet.A1�Z���I��"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      Application.Goto Reference:=Worksheets(sheetName).Range("A1"), Scroll:=True
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

'==================================================================================================
Function ���ׂĕ\��()
  Dim rowOutlineLevel As Long, colOutlineLevel As Long
  
  Const funcName As String = "Ctl_Sheet.���ׂĕ\��"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------

  If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
  End If
  If ActiveWindow.DisplayOutline = True Then
    ActiveSheet.Cells.ClearOutline
  End If
  ActiveSheet.Cells.EntireColumn.Hidden = False
  ActiveSheet.Cells.EntireRow.Hidden = False
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function �W�����()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim SelectAddress, setZoomLevel, resetBgColor, setGgridLine
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
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
      Application.Goto Reference:=Range("A1"), Scroll:=True
      
      'RC�\�L����AQ�\�L�֕ύX
      If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
      End If
      
    End If
    Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, sheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
'  Range(SelectAddress).Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
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
Function �V�[�g�Ǘ�_�t�H�[���\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim addSheetName As String
  Const funcName As String = "Ctl_Sheet.�V�[�g�Ǘ�_�t�H�[���\��"
  
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
  '----------------------------------------------
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���ݒ�()
  Dim SelectionCell As String

  
  Const funcName As String = "Ctl_Sheet.���ݒ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  Cells.Select
  Range("A1").Activate
  Selection.ColumnWidth = Library.getRegistry("Main", "ColumnWidth")
  
  Range(SelectionCell).Select
  
  '�����I��--------------------------------------
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
Function �����ݒ�()
  Dim SelectionCell As String

  Const funcName As String = "Ctl_Sheet.�����ݒ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  Cells.Select
  Range("A1").Activate
  Selection.rowHeight = Library.getRegistry("Main", "rowHeight")
  
  Range(SelectionCell).Select

  '�����I��--------------------------------------
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
Function �̍وꊇ�ύX()
  Dim setGgridLine As String
  
  Const funcName As String = "Ctl_Sheet.�̍وꊇ�ύX"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  '�I�[�g�V�F�C�v���̕���------------------------
  Dim slctObect As Shape
  On Error Resume Next
  For Each slctObect In ActiveSheet.Shapes
    slctObect.Select
    slctObect.Placement = xlMove
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = LadexsetVal("BaseFont")
      .TextRange.Font.NameFarEast = LadexsetVal("BaseFont")
      .TextRange.Font.Name = LadexsetVal("BaseFont")
      .TextRange.Font.Size = LadexsetVal("BaseFontSize")
      
      If .TextRange.Text <> "" Then
        .AutoSize = msoAutoSizeShapeToFitText
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        .WordWrap = msoTrue
      End If
    End With
    slctObect.Placement = xlFreeFloating
  Next
  On Error GoTo catchError
  
  '�K�C�h���C�����\��--------------------------
  setGgridLine = LadexsetVal("GridLine")
  If setGgridLine = "�\�����Ȃ�" Then
    ActiveWindow.DisplayGridlines = False
  ElseIf setGgridLine = "�\������" Then
    ActiveWindow.DisplayGridlines = True
  End If
      
  '�����ݒ�--------------------------------------
  Cells.Select
  Range("A1").Activate
  Selection.rowHeight = LadexsetVal("rowHeight")
  
  '�t�H���g�ݒ�----------------------------------
  With Selection
    .Font.Name = LadexsetVal("BaseFont")
    .Font.Size = LadexsetVal("BaseFontSize")
    .VerticalAlignment = xlCenter
  End With
  
  '�\���{��--------------------------------------
  ActiveWindow.Zoom = LadexsetVal("ZoomLevel")
  
  Application.Goto Reference:=Range("A1"), Scroll:=True

  '�����I��--------------------------------------
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
Function �w��t�H���g�ɐݒ�()
  Dim SelectionCell As String

  Const funcName As String = "Ctl_Sheet.�w��t�H���g�ɐݒ�"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  SelectionCell = Selection.Address
  
  '�I�[�g�V�F�C�v���̕���------------------------
  Dim slctObect As Shape
  On Error Resume Next
  For Each slctObect In ActiveSheet.Shapes
    slctObect.Select
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = LadexsetVal("BaseFont")
      .TextRange.Font.NameFarEast = LadexsetVal("BaseFont")
      .TextRange.Font.Name = LadexsetVal("BaseFont")
      .TextRange.Font.Size = LadexsetVal("BaseFontSize")
      If .TextRange.Text <> "" Then
        .AutoSize = msoAutoSizeShapeToFitText
        .AutoSize = msoAutoSizeNone
      End If

    End With
  
  
  Next
  On Error GoTo catchError
  
  '�t�H���g�ݒ�----------------------------------
  Cells.Select
  With Selection
    .Font.Name = LadexsetVal("BaseFont")
    .Font.Size = LadexsetVal("BaseFontSize")
    .VerticalAlignment = xlCenter
  End With

  Range(SelectionCell).Select
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function


'==================================================================================================
Function �A���V�[�g�ǉ�()
  Dim sheetName As Variant
  
  Const funcName As String = "Ctl_Sheet.�A���V�[�g�ǉ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
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
    
    Application.Goto Reference:=Range("A1"), Scroll:=True
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
Function ��ʘg�Œ�()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  Dim slctCell As String
  Const funcName As String = "Ctl_Sheet.��ʘg�Œ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  slctCell = ActiveCell.Address(False, False)
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    If Worksheets(sheetName).Visible = True And Not (sheetName Like "�s*�t") Then
      Call Library.showDebugForm("sheetName", sheetName, "debug")
      
      ActiveWindow.FreezePanes = False
      Range(slctCell).Select
      ActiveWindow.FreezePanes = True
      
      Application.Goto Reference:=Worksheets(sheetName).Range("A1"), Scroll:=True
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

