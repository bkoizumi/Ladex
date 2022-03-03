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

'**************************************************************************************************
' * �Z�����E��������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �Z��������()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  Dim slctCells As Range
  Const funcName As String = "Ctl_Sheet.�Z��������"
  
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
  If Selection.count > 1 Then
    Columns(Library.getColumnName(Selection(1).Column) & ":" & Library.getColumnName(Selection(Selection.count).Column)).EntireColumn.AutoFit
    
    For Each slctCells In Selection
      colName = Library.getColumnName(slctCells.Column)
      If Columns(colName & ":" & colName).ColumnWidth > 30 Then
        Columns(colName & ":" & colName).ColumnWidth = 30
      End If
    Next

  Else
    Cells.EntireColumn.AutoFit
    For colLine = 1 To Columns.count
      colName = Library.getColumnName(colLine)
      If IsNumeric(Cells(1, colLine)) Then
        If CInt(Cells(1, colLine)) > 1 Then
          Columns(colName & ":" & colName).ColumnWidth = Cells(1, colLine).Value
        End If
      End If
      
      If Cells(1, colLine).ColumnWidth > 30 Then
        colName = Library.getColumnName(colLine)
        Columns(colName & ":" & colName).ColumnWidth = 30
      End If
    Next
  End If
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript(True)
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function �Z����������()
  Const funcName As String = "Ctl_Sheet.�Z����������"
  
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
  
  Cells.EntireRow.AutoFit
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript(True)
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function A1�Z���I��()
  Dim objSheet As Object
  Dim SheetName As String, SetActiveSheet As String
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
    SheetName = objSheet.Name
    If Worksheets(SheetName).Visible = True Then
      Call Library.showDebugForm("sheetName", SheetName, "debug")
      Application.GoTo Reference:=Worksheets(SheetName).Range("A1"), Scroll:=True
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, sheetCount + 1, sheetMaxCount + 1, SheetName & "A1�Z���I��")
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
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
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
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
  Dim SheetName As String, SetActiveSheet As String
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
    SheetName = objSheet.Name
    If Worksheets(SheetName).Visible = True Then
      Worksheets(SheetName).Select
      
      '�W����ʂɐݒ�
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.View = xlNormalView
      
      '�\���{���̎w��
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      '�K�C�h���C���̕\��/��\��
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
      ActiveWindow.DisplayGridlines = setGgridLine
  
      '�w�i�����Ȃ��ɂ���
      Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
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
      'Application.Goto Reference:=Range("A1"), Scroll:=True
    End If
    Call Ctl_ProgressBar.showBar("�W����ʐݒ�", PrgP_Cnt, PrgP_Max, sheetCount, sheetMaxCount, SheetName)
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  Range(SelectAddress).Select
  
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
Function �V�[�g�Ǘ�_�t�H�[���\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim addSheetName As String
  Const funcName As String = "Ctl_Sheet.�V�[�g�Ǘ�_�t�H�[���\��"
  
  Set targetBook = ActiveWorkbook
  Frm_Sheet.Show vbModeless
  
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

