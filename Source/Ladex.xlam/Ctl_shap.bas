Attribute VB_Name = "Ctl_shap"
Option Explicit


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @Link https://infoment.hatenablog.com/entry/2021/08/17/000649
'**************************************************************************************************
Function TextToFitShape(targetShape As Excel.Shape, Optional chkFlg As Boolean = True) As Long
  ' �e�L�X�g�̗L���m�F�B�����ꍇ�́AFunction���I������B
  If targetShape.TextFrame2.TextRange.Characters.Text = vbNullString Then
      Exit Function
  End If

  ' �I�[�g�V�F�C�v�̃T�C�Y�擾�B
  Dim h(1) As Double: h(0) = targetShape.Height
  Dim w(1) As Double: w(0) = targetShape.Width
  Dim l As Double: l = targetShape.Left
  Dim T As Double: T = targetShape.Top
  
  ' �I�[�g�V�F�C�v����U�A�����T�C�Y�ɍ��킹�ăT�C�Y�ύX�B
  targetShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
  
  ' �ύX��̃T�C�Y�擾�B
  h(1) = targetShape.Height
  w(1) = targetShape.Width
  
  ' �I�[�g�V�F�C�v�̏c�Ɖ��A�e�X�̏k���i�������͊g��j���̂����A
  ' �����������擾�i�傫�������ƁA�I�[�g�V�F�C�v����H�ݏo��j�B
  Dim �� As Double
  �� = WorksheetFunction.Min(h(0) / h(1), w(0) / w(1))
  
  ' ���Ƃ̃t�H���g�T�C�Y�Ƀς��|���A�ڈ��̃t�H���g�T�C�Y�𓾂�B
  Dim FontSize As Long
  FontSize = targetShape.TextFrame2.TextRange.Font.Size * ��
      
  Dim i As Long
  Do
    ' �t�H���g�T�C�Y�����߁B
    targetShape.TextFrame2.TextRange.Font.Size = FontSize
    
    ' ���߂āA�I�[�g�V�F�C�v�𕶎��T�C�Y�ɍ��킹�ăT�C�Y�ύX�B
    targetShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    
    ' �ύX��̃T�C�Y�𓾂�B
    h(1) = targetShape.Height
    w(1) = targetShape.Width
    
    ' �c�Ɖ��ǂ��炩����ł����̃T�C�Y���z������A�����ŏI���B
    If (h(1) > h(0) Or w(1) > w(0)) And chkFlg = True Then
      Exit Do
    
    ElseIf (w(1) > w(0)) And chkFlg = False Then
      Exit Do
    
    ' �����łȂ���΁A�܂��s�b�^���ł͂Ȃ��B�t�H���g�T�C�Y���P�����B
    Else
        FontSize = FontSize + 1
    End If
    
    ' �������[�v�h�~�B
    i = i + 1: If i >= 100 Then Exit Do
  Loop
  
  ' �T�C�Y���z���Ă��甲�����̂ŁA�P�����Ē��x�̃T�C�Y�ɂ���B
  FontSize = FontSize - 1
  
  ' �I�[�g�T�C�Y�����B
  targetShape.TextFrame2.AutoSize = msoAutoSizeNone
  
  ' �I�[�g�V�F�C�v���ŏ��̑傫���ɖ߂��B
  targetShape.Height = h(0)
  targetShape.Width = w(0)
  
  targetShape.Left = l
  targetShape.Top = T
  
  ' �t�H���g�T�C�Y���ŏI�l�ɕύX�B
  targetShape.TextFrame2.TextRange.Font.Size = FontSize
  
  ' �߂�l�Ƃ��ăt�H���g�T�C�Y��Ԃ��B
  TextToFitShape = FontSize
End Function



'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************



'==================================================================================================
Function QR�R�[�h����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range, targetCells As Range
  
  Dim chartAPIURL As String
  Dim QRCodeImgName As String
  Dim colSize As Long, colHeight As Long, colWidth As Long
  
  Const funcName As String = "Ctl_Shap.QR�R�[�h����"
  Const chartAPI = "https://chart.googleapis.com/chart?cht=qr&chld=l|1&"
  
  
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

  With Frm_mkQRCode
    .Show
  End With
  
  
  For Each slctCells In Selection
    QRCodeImgName = "QRCode_" & slctCells.Address(False, False)
    
    '�������폜
    If Library.chkShapeName(QRCodeImgName) Then
      ActiveSheet.Shapes.Range(Array(QRCodeImgName)).Select
      Selection.delete
    End If
    
    colHeight = FrmVal("codeSize") * 0.75 + 4
    colWidth = FrmVal("codeSize") * 0.118 + 4
    Set targetCells = Range(FrmVal("CellAddress") & slctCells.Row)
    
    With targetCells
      .Select
      If FrmVal("onReSize") = True Then
        If .rowHeight < colHeight Then .rowHeight = colHeight
        If .ColumnWidth < colWidth Then .ColumnWidth = colWidth
      End If
    End With
    
    chartAPIURL = chartAPI & "chs=" & FrmVal("codeSize") & "x" & FrmVal("codeSize")
    chartAPIURL = chartAPIURL & "&chl=" & Library.convURLEncode(slctCells.Text)
    
    Call Library.showDebugForm("chartAPIURL", chartAPIURL, "debug")
    
    With ActiveSheet.Pictures.Insert(chartAPIURL)
      If FrmVal("onReSize") = True Then
        .ShapeRange.Top = targetCells.Top + (targetCells.Height - .ShapeRange.Height) / 2
        .ShapeRange.Left = targetCells.Left + (targetCells.Width - .ShapeRange.Width) / 2
      Else
        .ShapeRange.Top = targetCells.Top
        .ShapeRange.Left = targetCells.Left
      End If
      
      .Placement = xlMove
      
      'QR�R�[�h�̖��O�ݒ�
      .ShapeRange.Name = QRCodeImgName
      .Name = QRCodeImgName
    
    End With
    DoEvents
    Set targetCells = Nothing
  Next
  


  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
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
Function �p�X���[�h����()
  Const funcName As String = "Ctl_Shap.�p�X���[�h����"


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

  Frm_mkPasswd.Show vbModeless


  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
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
