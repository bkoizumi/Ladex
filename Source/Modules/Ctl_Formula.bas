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
  
'  On Error GoTo catchError
  Call Library.startScript
  
  '�����̃t�@�C���폜
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  If ActiveCell.HasFormula = False Or BKcf_rbPressed = False Then
    Call Library.endScript
    Exit Function
  End If
  
  Call init.setting
  aryRange = getFormulaRange(ActiveCell)
  
  count = 1
  For Each formulaVals In aryRange
    confirmFormulaName = "confirmFormulaName_" & count
  
    Call �͈͑I��(formulaVals, confirmFormulaName)
    count = count + 1
  Next
  
  ActiveCell.Select


  Call Library.endScript
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
  'Call Library.showNotice(400, "", True)
End Function


'==================================================================================================
Function �͈͑I��(formulaVals As Variant, confirmFormulaName As String)

  If formulaVals.Worksheet.Name <> ActiveSheet.Name Then
    Exit Function
  End If

  With ActiveSheet.Range(formulaVals.Address(external:=False))
    ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, width:=.width, height:=.height).Select
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
  
  
End Function



'==================================================================================================
Sub GetCurPosition()

  Dim p        As POINTAPI 'API�p�ϐ�
  Dim Rng  As Range
  Dim objShp
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmFormulaName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  Call Library.waitTime(50)
  Call Library.startScript
  
  '�J�[�\���ʒu�擾
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.Y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.Y).Select
  End If
  Call Library.endScript
  Call Ctl_Formula.�����m�F


End Sub

'==================================================================================================
Function getFormulaRange(ByVal argRange As Range) As Range()
    Dim sFormula As String
    Dim aryRange() As Range
    Dim tRange As Range
    Dim ix As Long
    Dim i As Long
    Dim flgS As Boolean '�V���O���N�I�[�g����̎�True
    Dim flgD As Boolean '�_�u���N�I�[�g����̎�True
    Dim sSplit() As String
    Dim sTemp As String
  
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
End Function



'**************************************************************************************************
' * �����ҏW
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function formula01()
  Dim formulaVal As String
  
  'On Error GoTo catchError
  
  If ActiveCell.HasFormula = False Then
    Exit Function
  End If
  Call init.setting
  
  formulaVal = ActiveCell.Formula
  formulaVal = Replace(formulaVal, "=", "")
  formulaVal = Replace(formulaVal, vbCrLf, "")
  formulaVal = Replace(formulaVal, vbLf, "")
  formulaVal = Trim(formulaVal)
  
  formulaVal = "IFERROR(" & formulaVal & ","""")"
  
  ActiveCell.Formula = "=" & formulaVal
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Function

'==================================================================================================
Function formula02()
  Dim formulaVal As String
  
  'On Error GoTo catchError

  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Function


'==================================================================================================
Function formula03()
  Dim formulaVal As String
  
  'On Error GoTo catchError

  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Function







