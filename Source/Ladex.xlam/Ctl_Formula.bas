Attribute VB_Name = "Ctl_Formula"
Option Explicit

'**************************************************************************************************
' * �������̃Z���Q��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @Link https://excel-ubara.com/excelvba5/EXCELVBA258.html
'**************************************************************************************************
'==================================================================================================
Function �����m�F()
  Dim confirmFormulaName As String
  Dim count As Long
  Dim formulaVals As Variant
  Dim objShp, aryRange
  Const funcName As String = "Ctl_Formula.�����m�F"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  '�����̃I�u�W�F�N�g�폜
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  If ActiveCell.HasFormula = False Or BKcf_rbPressed = False Then
    Call Library.endScript
    Exit Function
  End If
  aryRange = Ctl_Formula.getFormulaRange(ActiveCell)
  
  count = 1
  For Each formulaVals In aryRange
    confirmFormulaName = "confirmFormulaName_" & count
    Call Ctl_Formula.�͈͑I��(formulaVals, confirmFormulaName)
    count = count + 1
  Next
  ActiveCell.Select
  
  '�����I��--------------------------------------
  Call Library.endScript
  If runFlg = False Then
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
Function �͈͑I��(formulaVals As Variant, confirmFormulaName As String)
  Const funcName As String = "Ctl_Formula.�͈͑I��"
  
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
  
  If formulaVals.Worksheet.Name <> ActiveSheet.Name Then
    Exit Function
  End If

  With ActiveSheet.Range(formulaVals.Address(external:=False))
    ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height).Select
  End With
  
  Selection.Name = confirmFormulaName
  Selection.ShapeRange.Fill.ForeColor.RGB = RGB(205, 205, 255)
  Selection.ShapeRange.Fill.Transparency = 0.5
  Selection.OnAction = "Ctl_Formula.GetCurPosition"
  Selection.Text = formulaVals.Address(RowAbsolute:=False, ColumnAbsolute:=False, external:=False)
  
  With Selection.ShapeRange.TextFrame2
    .TextRange.Font.NameComplexScript = "���C���I"
    .TextRange.Font.NameFarEast = "���C���I"
    .TextRange.Font.Name = "���C���I"
    .TextRange.Font.Size = 9
    .MarginLeft = 3
    .MarginRight = 0
    .MarginTop = 0
    .MarginBottom = 0
    .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
  End With
  
  Selection.ShapeRange.line.Visible = msoTrue
  Selection.ShapeRange.line.ForeColor.RGB = RGB(255, 0, 0)
  Selection.ShapeRange.line.Weight = 2
  
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
Function GetCurPosition()
  Dim p As POINTAPI 'API�p�ϐ�
  Dim Rng  As Range
  Dim objShp
  Const funcName As String = "Ctl_Formula.GetCurPosition"
  
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
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  Call Library.waitTime(50)
  
  '�J�[�\���ʒu�擾
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.y).Select
  End If
  Call Ctl_Formula.�����m�F
  
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
Function getFormulaRange(ByVal argRange As Range) As Range()
  Dim sFormula As String, sSplit() As String, sTemp As String
  Dim aryRange() As Range, tRange As Range
  Dim ix As Long, i As Long
  Dim flgS As Boolean, flgD As Boolean
  Const funcName As String = "Ctl_Formula.getFormulaRange"
  
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
  '=�ȍ~�̌v�Z��
  sFormula = Mid(argRange.FormulaLocal, 2)
  '�v�Z���̒��̉��s��]���ȋ󔒂�����
  sFormula = Replace(sFormula, vbCrLf, "")
  sFormula = Replace(sFormula, vbLf, "")
  sFormula = Trim(sFormula)

  flgS = False
  flgD = False
  For i = 1 To Len(sFormula)
    '�V���O���E�_�u����True,False�𔽓]
    Select Case Mid(sFormula, i, 1)
      Case "'"
        flgS = Not flgS
      Case """"
        '�V���O���̒��Ȃ�V�[�g��
        If Not flgS Then
          flgD = Not flgD
        End If
    End Select
    Select Case Mid(sFormula, i, 1)
      '�e�퉉�Z�q�̔���
      Case "+", "-", "*", "/", "^", ">", "<", "=", "(", ")", "&", ",", " "
        Select Case True
          Case flgS
            '�V���O���̒��Ȃ�V�[�g��
            sTemp = sTemp & Mid(sFormula, i, 1)
          Case flgD
            '�_�u���̒��Ȃ疳��
          Case Else
            '�e�퉉�Z�q��vbLf�ɒu��
            sTemp = sTemp & vbLf
        End Select
      Case Else
        '�_�u���̒��Ȃ疳���A�������V���O���̒��̓V�[�g��
        If Not flgD Or flgS Then
          sTemp = sTemp & Mid(sFormula, i, 1)
        End If
    End Select
  Next

  On Error Resume Next
  'vbLf�ŋ�؂��Ĕz��
  sSplit = Split(sTemp, vbLf)
  ix = 0
  For i = 0 To UBound(sSplit)
    If sSplit(i) <> "" Then
      Err.Clear
      'Application.Evaluate���\�b�h���g����Range�ɕϊ�
      If InStr(sSplit(i), "!") > 0 Then
        Set tRange = Evaluate(Trim(sSplit(i)))
      Else
        '�V�[�g�����܂܂Ȃ��ꍇ�́A���Z���̃V�[�g����t��
        Set tRange = Evaluate("'" & argRange.Parent.Name & "'!" & Trim(sSplit(i)))
      End If
      'Range�I�u�W�F�N�g������������Δz��֓����
      If Err.Number = 0 Then
        ReDim Preserve aryRange(ix)
        Set aryRange(ix) = tRange
        ix = ix + 1
      End If
    End If
  Next
  On Error GoTo 0
  getFormulaRange = aryRange
  
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

'**************************************************************************************************
' * �����ҏW
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �G���[�h�~_��()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.�G���[�h�~_��"
  
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
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IFERROR(" & formulaVal & ","""")"
      
      slctCells.Formula = "=" & formulaVal
    End If
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
Function �G���[�h�~_�[��()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.�G���[�h�~_�[��"
  
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
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IFERROR(" & formulaVal & ",0)"
      
      slctCells.Formula = "=" & formulaVal
    End If
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
Function �[����\��()
  Dim slctCells As Range
  Dim formulaVal As String
  Const funcName As String = "Ctl_Formula.�[����\��"
  
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
  
  
  For Each slctCells In Selection
    If slctCells.HasFormula = True Then
      formulaVal = slctCells.Formula
      formulaVal = Replace(formulaVal, "=", "")
      formulaVal = Replace(formulaVal, vbCrLf, "")
      formulaVal = Replace(formulaVal, vbLf, "")
      formulaVal = Trim(formulaVal)
      
      formulaVal = "IF(" & formulaVal & "=0,""""," & formulaVal & ")"
      
      slctCells.Formula = "=" & formulaVal
    End If
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
Function �s�ԍ��ǉ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Formula.�s�ԍ��ǉ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  line = Selection.Row - 1
  For Each slctCells In Selection
    slctCells.FormulaR1C1 = "=ROW()-" & line
    DoEvents
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
Function �����}��(formulaType As String)
  Dim confirmFormulaName As String
  Dim count As Long
  Dim formulaVals As Variant
  Dim objShp, aryRange
  Const funcName As String = "Ctl_Formula.�����}��"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.showDebugForm("formulaType", formulaType, "debug")
  '----------------------------------------------
  
  Select Case formulaType
    Case "SheetName"
      ActiveCell.FormulaR1C1 = "=MID(CELL(""filename"",RC),FIND(""]"",CELL(""filename"",RC))+1,31)"


    Case Else
  End Select

  
  '�����I��--------------------------------------
  Call Library.endScript
  If runFlg = False Then
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
