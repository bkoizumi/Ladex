Attribute VB_Name = "Ctl_Chart"

'**************************************************************************************************
' * �_����WBS�p�`���[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �_����WBS�p�`���[�g����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Chart.�_����WBS�p�`���[�g����"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
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
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'**************************************************************************************************
' * �K���g�`���[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �K���g�`���[�g����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim startColumn As String, endColumn As String
  
  Call WBS_Option.�I���V�[�g�m�F
  
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call �K���g�`���[�g�폜
  endLine = Cells(Rows.count, 2).End(xlUp).Row
  
  For line = 6 To endLine
    '�v�������------------------------------------
    If Not (mainSheet.Range(setVal("GUNT_START_DAY") & line) = "" Or mainSheet.Range(setVal("GUNT_END_DAY") & line) = "") Then
      Call �v����ݒ�(line)
    End If

    '���ѐ�����------------------------------------
    If mainSheet.Range(setVal("cell_Progress") & line) >= 0 Then
      Call ���ѐ��ݒ�(line)
    End If
    
    '�^�C�����C���ւ̒ǉ�------------------------------------
    If (mainSheet.Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")) Then
      Call �^�C�����C���ɒǉ�(line)
    End If
    
    '�C�i�Y�}������------------------------------
    Call �C�i�Y�}���ݒ�(line)
    
    '�i����100%�Ȃ��\��------------------------------------
    If setVal("setDispProgress100") = True And mainSheet.Range(setVal("cell_Progress") & line) = 100 Then
      Rows(line & ":" & line).EntireRow.Hidden = True
      
    End If

  Next
  For line = 6 To endLine
    Call �^�X�N�̃����N�ݒ�(line)
  Next

  If ActiveSheet.Name = mainSheetName Then
    Call WBS_Option.�����̒S���ҍs���\��
  End If

End Function


'**************************************************************************************************
' * �K���g�`���[�g�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �K���g�`���[�g�폜()
  Dim shp As Shape
  Dim rng As Range
  
  On Error Resume Next
  
  Set rng = Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(Rows.count, Columns.count))
  
  For Each shp In ActiveSheet.Shapes
    If Not Intersect(Range(shp.TopLeftCell, shp.BottomRightCell), rng) Is Nothing Then
      If (shp.Name Like "Drop Down*") Or (shp.Name Like "Comment*") Then
      Else
        'Debug.Print shp.Name
        shp.Select
        shp.Delete
      End If
    End If
  Next

End Function


'**************************************************************************************************
' * �v����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �v����ݒ�(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  
  startColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_START_DAY") & line))
  endColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line))
  
'  'Shape��z�u���邽�߂̊�ƂȂ�Z��
'  Set rngStart = mainSheet.Range(startColumn & line)
'  Set rngEnd = mainSheet.Range(endColumn & line)
'
'  '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
'  BX = rngStart.Left
'  BY = rngStart.top + (rngStart.Height / 2)
'  EX = rngEnd.Left + rngEnd.Width
'  EY = rngEnd.top + (rngEnd.Height / 2)
  
  '�S���ҕʂ̐F�ݒ�------------------------------
  lColorValue = 0
  If Range(setVal("cell_Assign") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_Assign") & line).Value)
  ElseIf Range(setVal("cell_Assign") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_Assign") & line).Value)
  End If
  
  
  
  If lColorValue <> 0 And ActiveSheet.Name = mainSheetName Then
    Call Library.getRGB(lColorValue, Red, Green, Blue)
  Else
    Call Library.getRGB(setVal("lineColor_Plan"), Red, Green, Blue)
  End If

  If Range(setVal("cell_Assign") & line) = "�H��" Or Range(setVal("cell_Assign") & line) = "�H��" Then
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "�^�X�N_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
'        .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
'        .TextFrame.Characters.Font.Size = 12
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
      End With
    End With
    Set ProcessShape = Nothing
    ActiveSheet.Shapes.Range(Array("�^�X�N_" & line)).Select
    Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
  
  Else
    With Range(startColumn & line & ":" & endColumn & line)
      'Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "�^�X�N_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
        
        '.TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
        .TextFrame.Characters.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
    
        If setVal("viewGant_TaskName") = True Then
          ActiveSheet.Shapes.Range(Array("�^�X�N_" & line)).Select
          Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
        End If
        
        .OnAction = "beforeChangeShapes"
      End With
    End With
    Set ProcessShape = Nothing

    '�S���Җ���\��
    If setVal("viewGant_Assignor") = True Then
      startColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line) + 1)
      endColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line) + 3)

      With Range(startColumn & line & ":" & endColumn & line)
        Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left + 10, Top:=.Top, Width:=.Width + 10, Height:=10)
        
        With ProcessShape
          .Name = "�S����_" & line
          .Fill.ForeColor.RGB = RGB(255, 255, 255)
          .Fill.Transparency = 0
          '.TextFrame.Characters.Text = Range(setVal("cell_Assign") & line)
          .TextFrame.Characters.Font.Size = 9
          .TextFrame2.TextRange.Font.Bold = msoTrue
          .TextFrame2.WordWrap = msoFalse
          .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
          .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
          .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
          .TextFrame2.VerticalAnchor = msoAnchorMiddle
          .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
          .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
          .TextFrame2.TextRange.Font.Name = "���C���I"
          .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
          .TextFrame2.TextRange.Font.Size = 9
        End With
      End With
      Set ProcessShape = Nothing
      ActiveSheet.Shapes.Range(Array("�S����_" & line)).Select
      Selection.Formula = "=" & Range(setVal("cell_Assign") & line).Address(False, False)
    End If
  End If
End Function


'**************************************************************************************************
' * ���ѐ��ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���ѐ��ݒ�(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  Dim shapesWith As Long
  
'    lColorValue = setSheet.Range(setVal("cell_ProgressEnd") & line).Interior.Color
  
'  Call Library.showDebugForm("���ѐ��ݒ�", Range(setVal("cell_TaskArea") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�J�n��:" & Range(setVal("cell_AchievementStart") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�I����:" & Range(setVal("cell_AchievementEnd") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�i���@:" & Range(setVal("cell_Progress") & line))
  
  If Range(setVal("cell_AchievementStart") & line) = "" Then
    startColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_START_DAY") & line))
  Else
    startColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementStart") & line))
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) = "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line))
  
  '�i����100%�̂Ƃ�
  ElseIf Range(setVal("cell_Progress") & line) = 100 Then
    If Range(setVal("cell_AchievementEnd") & line) < Range(setVal("GUNT_END_DAY") & line) Then
      endColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line))
    Else
      endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
    End If
  
  Else
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
  End If

  
  
  Call Library.getRGB(setVal("lineColor_Achievement"), Red, Green, Blue)

  
  With Range(startColumn & line & ":" & endColumn & line)
    .Select
    
    If Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_Progress") & line) = 0 Then
      shapesWith = 0
    Else
      shapesWith = .Width * (Range(setVal("cell_Progress") & line) / 100)
    End If
    
    If Range(setVal("cell_Assign") & line) = "�H��" Or Range(setVal("cell_Assign") & line) = "�H��" Then
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, Top:=.Top + 5, Width:=shapesWith, Height:=10)
    Else
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top + 5, Width:=shapesWith, Height:=10)
    End If
    
    With ProcessShape
      .Name = "����_" & line
      .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
      .Fill.Transparency = 0.6
    End With
  End With
  Set ProcessShape = Nothing
    
    
    


End Function


'**************************************************************************************************
' * �C�i�Y�}���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �C�i�Y�}���ݒ�(line As Long)

  Dim rngStart As Range, rngEnd As Range, rngBase As Range, rngChkDay As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim startColumn As String, endColumn As String, baseColumn As String, chkDayColumn As String
  Dim progress As Long, lateOrEarly As Double
  Dim extensionDay As Integer
  Dim chkDay As Date
  Dim Red As Long, Green As Long, Blue As Long
  
  If Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("GUNT_END_DAY")) Then
    If setVal("setLightning") = True Then
      Call Library.showNotice(50)
      setVal("setLightning") = False
      Range("setLightning") = False
    End If
    Exit Function
    
  End If
  
  '�C�i�Y�}���̐F�擾
  Call Library.getRGB(setVal("lineColor_Lightning"), Red, Green, Blue)
  
  baseColumn = WBS_Option.���t�Z������(setVal("baseDay"))
  
  '�^�C�����C����Ɉ���
  If line = 6 Then
    Set rngBase = Range(baseColumn & 5)
    
    '�����R�l�N�^����
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��B_5"
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
  End If
  
  Set rngBase = Range(baseColumn & line)
  
  
  
  
  '�C�i�Y�}���������Ȃ��ꍇ�́A����݈̂���
  If setVal("setLightning") = False Or Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_LateOrEarly") & line) = 0 Then
    
    '�����R�l�N�^����
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��B_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
    Exit Function
  
  '�i����0%�ȏ�̏ꍇ�́A�C�i�Y�}��������
  ElseIf Range(setVal("cell_Progress") & line) >= 0 Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("����_" & line), 4
  
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes("����_" & line), 4
    
    
'
'      startTask = "�^�X�N_" & tmpLine
'      thisTask = "�^�X�N_" & line
'
'    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
'    Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
'
'    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
'    Selection.Name = "�C�i�Y�}��_" & line
  
      
  End If
Exit Function





  If Range(setVal("GUNT_START_DAY") & line) <> "" Then
    startColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_START_DAY") & line))
  Else
    startColumn = baseColumn
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
  
  ElseIf Range(setVal("GUNT_END_DAY") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("GUNT_END_DAY") & line))
  Else
    endColumn = baseColumn
  End If
    
  'Shape��z�u���邽�߂̊�ƂȂ�Z��
  Set rngStart = Range(startColumn & line)
  Set rngEnd = Range(endColumn & line)

  
  '�x���H���̒l
  If Range(setVal("cell_LateOrEarly") & line) = 0 Or Range(setVal("cell_LateOrEarly") & line) = "" Then
    BX = rngBase.Left + rngBase.Width
    BY = rngBase.Top
    EX = rngBase.Left + rngBase.Width
    EY = rngBase.Top + rngBase.Height
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With
  Else
    chkDay = WBS_Option.�C�i�Y�}���p���t�v�Z(setVal("baseDay"), Range(setVal("cell_LateOrEarly") & line))
    chkDayColumn = WBS_Option.���t�Z������(chkDay)
    
    Set rngChkDay = Range(chkDayColumn & line)
    
    BX = rngBase.Left + rngBase.Width
    BY = rngBase.Top
    EX = rngChkDay.Left + rngChkDay.Width
    EY = rngBase.Top + (rngBase.Height / 2)
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With

    BX = rngChkDay.Left + rngChkDay.Width
    BY = rngBase.Top + (rngBase.Height / 2)
    EX = rngBase.Left + rngBase.Width
    EY = rngBase.Top + rngBase.Height
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With
  End If
  
End Function


'**************************************************************************************************
' * �^�C�����C���ɒǉ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�C�����C���ɒǉ�(line As Long, Optional autoFlg As Boolean = False)
  Dim endLine As Long, colLine As Long, endColLine As Long
  Dim ShapeTopStart As Long, count As Long
  Dim shp As Shape
  Dim rng As Range
  Dim colorVal As Long
  Dim targetTaskName As String
  
  Const funcName As String = "Ctl_Chart.�^�C�����C���ɒǉ�"


'  On Error GoTo catchError
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "function1")
  
  startColumn = WBS_Option.���t�Z������(Format(Range("M" & line), "yyyy/mm/dd"))
  endColumn = WBS_Option.���t�Z������(Format(Range("N" & line), "yyyy/mm/dd"))


  If Library.chkShapeName("TimeLine_" & line) Then
    ActiveSheet.Shapes("TimeLine_" & line).Delete
  End If
  
  On Error Resume Next
  count = 0
  Range(startColumn & "6:" & endColumn & 6).Select
  For Each shp In ActiveSheet.Shapes
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
    If Not (Intersect(rng, Selection) Is Nothing) Then
      count = count + 1
    End If
  Next
  If count <> 0 Then
    ShapeTopStart = 15 * count
  Else
    ShapeTopStart = 0
  End If
  On Error GoTo 0



  '�^�C�����C���s�̕����L����
  If count >= 2 Then
    Rows("6:6").RowHeight = Rows("6:6").RowHeight + 15
  End If
  
  targetTaskName = Task.getTaskName(line)
  Call Library.showDebugForm("targetTaskName", targetTaskName, "debug")
  Call Library.showDebugForm("Color", Range("B" & line).Interior.Color, "debug")
  
  If Range("B" & line).Interior.Color = 16777215 Or Range("B" & line).Interior.Color = 2704713 Then
    colorVal = RGB(102, 102, 255)
  Else
    colorVal = Range("B" & line).Interior.Color
  End If
  
  If startColumn = endColumn And (Range("K" & line) = "���J" Or Range("L" & line) = "���J") Then
    endLine = Cells(Rows.count, 1).End(xlUp).Row
    With Range(startColumn & "6:" & endColumn & endLine)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "TimeLine_" & line
        .Fill.ForeColor.RGB = colorVal
        .Fill.Transparency = 0.6
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
      End With
    End With
    ActiveSheet.Shapes.Range(Array("TimeLine_" & line)).Select
'    Selection.Text = Task.getTaskName(line)
'    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
'    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
    
    With Selection
      .Text = targetTaskName
      .ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
      .ShapeRange.TextFrame2.VerticalAnchor = msoAnchorTop
      .ShapeRange.line.Visible = msoTrue
      .ShapeRange.line.Weight = 1.5
      .ShapeRange.line.ForeColor.RGB = colorVal
      .ShapeRange.line.Transparency = 0
      .ShapeRange.TextFrame2.MarginLeft = 0
      .ShapeRange.TextFrame2.MarginRight = 0
      .ShapeRange.TextFrame2.MarginTop = 0
      .ShapeRange.TextFrame2.MarginBottom = 0
    End With
    
    
    
  Else
    With Range(startColumn & "6:" & endColumn & 6)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left, Top:=.Top + ShapeTopStart, Width:=.Width, Height:=15)
      
      With ProcessShape
        .Name = "TimeLine_" & line
        .Fill.ForeColor.RGB = colorVal
        .Fill.Transparency = 0.6
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
      End With
    End With
    ActiveSheet.Shapes.Range(Array("TimeLine_" & line)).Select
    With Selection
      .Text = targetTaskName
      .ShapeRange.line.Visible = msoTrue
      .ShapeRange.line.Weight = 1.5
      .ShapeRange.line.ForeColor.RGB = colorVal
      .ShapeRange.line.Transparency = 0
      .ShapeRange.TextFrame2.MarginLeft = 0
      .ShapeRange.TextFrame2.MarginRight = 0
      .ShapeRange.TextFrame2.MarginTop = 0
      .ShapeRange.TextFrame2.MarginBottom = 0
    End With
  End If
  
  If autoFlg = False Then
    Sh_PARAM.Range("AL" & Sh_PARAM.Cells(Rows.count, 38).End(xlUp).Row + 1).Formula = "=ROW(WBS!" & Range("A" & line).Address(False, False) & ")"

    
    '�d���폜
    Sh_PARAM.Columns("AL:AL").RemoveDuplicates Columns:=1, Header:=xlNo
    
    Sh_PARAM.Sort.SortFields.Clear
    Sh_PARAM.Sort.SortFields.Add Key:=Range("AL2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PARAM").Sort
        .SetRange Range("AL2:AL" & Sh_PARAM.Cells(Rows.count, 38).End(xlUp).Row)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
  End If
  
'  If Range(setVal("cell_Info") & line) = "" Then
'    Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")
'  ElseIf Range(setVal("cell_Info") & line) Like "*" & setVal("TaskInfoStr_TimeLine") & "*" Then
'  Else
'    Range(setVal("cell_Info") & line) = Range(setVal("cell_Info") & line) & "," & setVal("TaskInfoStr_TimeLine")
'  End If

  Range(Task.getTaskName(line, "address")).Select

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �Z���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �Z���^�[()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseDate As Date
  Dim baseColumn As String
  
  
'  On Error GoTo catchError

  If setVal("startDay") >= setVal("baseDay") - 10 Then
    baseDate = setVal("startDay")
  Else
    baseDate = setVal("baseDay") - 10
    
  End If
  
  baseColumn = WBS_Option.���t�Z������(baseDate)
  Application.Goto Reference:=Range(baseColumn & 6), Scroll:=True


  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function





'**************************************************************************************************
' * �K���g�`���[�g�I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function beforeChangeShapes()

  Call Library.startScript
  ActiveSheet.Shapes.Range(Array(Application.Caller)).Select
  changeShapesName = Application.Caller
  
'  Call Library.setArrayPush(selectShapesName, Application.Caller)
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With
  
  Call Library.endScript
End Function


'**************************************************************************************************
' * Chart����\����𑀍�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function changeShapes()
  Dim rng As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
  
  Call Library.showDebugForm("Chart����\����𑀍�", "�����J�n")
  
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With

  With ActiveSheet.Shapes(changeShapesName)
    Set rng = Range(.TopLeftCell, .BottomRightCell)
  End With
  If rng.Address(False, False) Like "*:*" Then
    tmp = Split(rng.Address(False, False), ":")
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName) = Range(getColumnName(Range(tmp(0)).Column) & 4)
    mainSheet.Range(setVal("GUNT_END_DAY") & changeShapesName) = Range(getColumnName(Range(tmp(1)).Column) & 4)
  Else
    tmp = rng.Address(False, False)
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
    mainSheet.Range(setVal("GUNT_END_DAY") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
  End If
  
  '��s�^�X�N�̏I����+1���J�n���ɐݒ�
  newStartDay = mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName)
  Call init.chkHollyday(newStartDay, HollydayName)
  Do While HollydayName <> ""
    newStartDay = newStartDay - 1
    Call init.chkHollyday(newStartDay, HollydayName)
  Loop
  Range(setVal("GUNT_START_DAY") & changeShapesName) = newStartDay
  
  '�I�������Đݒ�
  newEndDay = mainSheet.Range(setVal("GUNT_END_DAY") & changeShapesName)
  Call init.chkHollyday(newEndDay, HollydayName)
  Do While HollydayName <> ""
    newEndDay = newEndDay + 1
    Call init.chkHollyday(newEndDay, HollydayName)
  Loop
  Range(setVal("GUNT_END_DAY") & changeShapesName) = newEndDay
  
  If ActiveSheet.Name = TeamsPlannerSheetName Then
    If Range(setVal("cell_Info") & changeShapesName) = "" Then
      Range(setVal("cell_Info") & changeShapesName) = setVal("TaskInfoStr_Change")
    ElseIf Range(setVal("cell_Info") & changeShapesName) Like "*" & setVal("TaskInfoStr_Change") & "*" Then
    Else
      Range(setVal("cell_Info") & changeShapesName) = Range(setVal("cell_Info") & changeShapesName) & "," & setVal("TaskInfoStr_Change")
    End If
  End If
  
  ActiveSheet.Shapes("�^�X�N_" & changeShapesName).Delete
  If Range(setVal("cell_Task") & CLng(changeShapesName)) <> "" Then
    ActiveSheet.Shapes("��s�^�X�N�ݒ�_" & CLng(changeShapesName)).Delete
  End If
  If setVal("viewGant_Assignor") = True Then
    ActiveSheet.Shapes("�S����_" & changeShapesName).Delete
  End If
  
  Call �v����ݒ�(CLng(changeShapesName))
  Call �^�X�N�̃����N�ݒ�(CLng(changeShapesName))
  
  changeShapesName = ""

  Call Library.showDebugForm("Chart����\����𑀍�", "�����I��")
End Function


'**************************************************************************************************
' * �^�X�N�̃����N�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�̃����N�ݒ�(line As Long)
  Dim startTask As String, thisTask As String
  Dim interval As Long
  Dim tmpLine As Variant
  
'  On Error GoTo catchError


  If Range(setVal("cell_Task") & line) = "" Then
    Exit Function
  Else
    For Each tmpLine In Split(Range(setVal("cell_Task") & line), ",")
      startTask = "�^�X�N_" & tmpLine
      thisTask = "�^�X�N_" & line
    
      '�J�M���R�l�N�^����
      ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
      Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
      Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startTask), 4
      Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
      Selection.Name = "��s�^�X�N�ݒ�_" & line
      
      interval = DateDiff("d", Range(setVal("GUNT_END_DAY") & tmpLine), Range(setVal("GUNT_START_DAY") & line))
      If interval < 1 Then
        interval = interval * -1
      End If
      
     If interval = 0 Then
     ElseIf interval < 2 Then
        Selection.ShapeRange.Flip msoFlipHorizontal
     End If
      
      With Selection.ShapeRange.line
        .Visible = msoTrue
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
      End With
    Next
  End If


  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function �@�\�ǉ�_�v���()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim shp As Shape
  
  
  
  Const funcName As String = "Ctl_Chart.�@�\�ǉ�_�v���"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  For Each shp In ActiveSheet.Shapes
    Select Case True
      Case shp.Name Like "CommandButton*"
      Case shp.Name Like "TextBox*"
      Case Else
        line = shp.TopLeftCell.Row
        
        'B��Ɠ����F�Ȃ�v����ƔF�����ď���
        If shp.line.ForeColor = Range("B" & line).Interior.Color Or shp.line.ForeColor = 65370 Then
          shp.Name = "taskLine_" & line
'          shp.OnAction = "beforeChangeShapes"
        End If
        
        Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "")
        
    End Select
    
  Next
  

  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
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
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function







