Attribute VB_Name = "Ctl_Chart"

'**************************************************************************************************
' * 濱さんWBS用チャート生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 濱さんWBS用チャート生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Chart.濱さんWBS用チャート生成"

  '処理開始--------------------------------------
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
  
  '濱さん作成VBAの呼び出し
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
  
  'タイムラインに追加----------------------------
  Call Library.startScript
  Rows("6:6").RowHeight = 40
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "TimeLine_*" Then
      ActiveSheet.Shapes(objShp.Name).Delete
    End If
  Next
  
  
  For line = 2 To Sh_PARAM.Cells(Rows.count, 38).End(xlUp).Row
    If Sh_PARAM.Range("AL" & line).Text <> "" Then
      Call Ctl_Chart.タイムラインに追加(CLng(Sh_PARAM.Range("AL" & line).Text), True)
    End If
  Next
  '処理終了--------------------------------------
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

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'**************************************************************************************************
' * ガントチャート生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ガントチャート生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim startColumn As String, endColumn As String
  
  Call WBS_Option.選択シート確認
  
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call ガントチャート削除
  endLine = Cells(Rows.count, 2).End(xlUp).Row
  
  For line = 6 To endLine
    '計画線生成------------------------------------
    If Not (mainSheet.Range(setVal("GUNT_START_DAY") & line) = "" Or mainSheet.Range(setVal("GUNT_END_DAY") & line) = "") Then
      Call 計画線設定(line)
    End If

    '実績線生成------------------------------------
    If mainSheet.Range(setVal("cell_Progress") & line) >= 0 Then
      Call 実績線設定(line)
    End If
    
    'タイムラインへの追加------------------------------------
    If (mainSheet.Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")) Then
      Call タイムラインに追加(line)
    End If
    
    'イナズマ線生成------------------------------
    Call イナズマ線設定(line)
    
    '進捗が100%なら非表示------------------------------------
    If setVal("setDispProgress100") = True And mainSheet.Range(setVal("cell_Progress") & line) = 100 Then
      Rows(line & ":" & line).EntireRow.Hidden = True
      
    End If

  Next
  For line = 6 To endLine
    Call タスクのリンク設定(line)
  Next

  If ActiveSheet.Name = mainSheetName Then
    Call WBS_Option.複数の担当者行を非表示
  End If

End Function


'**************************************************************************************************
' * ガントチャート削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ガントチャート削除()
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
' * 計画線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 計画線設定(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  
  startColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_START_DAY") & line))
  endColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line))
  
'  'Shapeを配置するための基準となるセル
'  Set rngStart = mainSheet.Range(startColumn & line)
'  Set rngEnd = mainSheet.Range(endColumn & line)
'
'  'セルのLeft、Top、Widthプロパティを利用して位置決め
'  BX = rngStart.Left
'  BY = rngStart.top + (rngStart.Height / 2)
'  EX = rngEnd.Left + rngEnd.Width
'  EY = rngEnd.top + (rngEnd.Height / 2)
  
  '担当者別の色設定------------------------------
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

  If Range(setVal("cell_Assign") & line) = "工程" Or Range(setVal("cell_Assign") & line) = "工程" Then
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "タスク_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
'        .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
'        .TextFrame.Characters.Font.Size = 12
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
      End With
    End With
    Set ProcessShape = Nothing
    ActiveSheet.Shapes.Range(Array("タスク_" & line)).Select
    Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
  
  Else
    With Range(startColumn & line & ":" & endColumn & line)
      'Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "タスク_" & line
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
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
    
        If setVal("viewGant_TaskName") = True Then
          ActiveSheet.Shapes.Range(Array("タスク_" & line)).Select
          Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
        End If
        
        .OnAction = "beforeChangeShapes"
      End With
    End With
    Set ProcessShape = Nothing

    '担当者名を表示
    If setVal("viewGant_Assignor") = True Then
      startColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line) + 1)
      endColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line) + 3)

      With Range(startColumn & line & ":" & endColumn & line)
        Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left + 10, Top:=.Top, Width:=.Width + 10, Height:=10)
        
        With ProcessShape
          .Name = "担当者_" & line
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
          .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
          .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
          .TextFrame2.TextRange.Font.Name = "メイリオ"
          .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
          .TextFrame2.TextRange.Font.Size = 9
        End With
      End With
      Set ProcessShape = Nothing
      ActiveSheet.Shapes.Range(Array("担当者_" & line)).Select
      Selection.Formula = "=" & Range(setVal("cell_Assign") & line).Address(False, False)
    End If
  End If
End Function


'**************************************************************************************************
' * 実績線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 実績線設定(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  Dim shapesWith As Long
  
'    lColorValue = setSheet.Range(setVal("cell_ProgressEnd") & line).Interior.Color
  
'  Call Library.showDebugForm("実績線設定", Range(setVal("cell_TaskArea") & line))
'  Call Library.showDebugForm("実績線設定", "　開始日:" & Range(setVal("cell_AchievementStart") & line))
'  Call Library.showDebugForm("実績線設定", "　終了日:" & Range(setVal("cell_AchievementEnd") & line))
'  Call Library.showDebugForm("実績線設定", "　進捗　:" & Range(setVal("cell_Progress") & line))
  
  If Range(setVal("cell_AchievementStart") & line) = "" Then
    startColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_START_DAY") & line))
  Else
    startColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementStart") & line))
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) = "" Then
    endColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line))
  
  '進捗が100%のとき
  ElseIf Range(setVal("cell_Progress") & line) = 100 Then
    If Range(setVal("cell_AchievementEnd") & line) < Range(setVal("GUNT_END_DAY") & line) Then
      endColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line))
    Else
      endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
    End If
  
  Else
    endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
  End If

  
  
  Call Library.getRGB(setVal("lineColor_Achievement"), Red, Green, Blue)

  
  With Range(startColumn & line & ":" & endColumn & line)
    .Select
    
    If Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_Progress") & line) = 0 Then
      shapesWith = 0
    Else
      shapesWith = .Width * (Range(setVal("cell_Progress") & line) / 100)
    End If
    
    If Range(setVal("cell_Assign") & line) = "工程" Or Range(setVal("cell_Assign") & line) = "工程" Then
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, Top:=.Top + 5, Width:=shapesWith, Height:=10)
    Else
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top + 5, Width:=shapesWith, Height:=10)
    End If
    
    With ProcessShape
      .Name = "実績_" & line
      .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
      .Fill.Transparency = 0.6
    End With
  End With
  Set ProcessShape = Nothing
    
    
    


End Function


'**************************************************************************************************
' * イナズマ線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function イナズマ線設定(line As Long)

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
  
  'イナズマ線の色取得
  Call Library.getRGB(setVal("lineColor_Lightning"), Red, Green, Blue)
  
  baseColumn = WBS_Option.日付セル検索(setVal("baseDay"))
  
  'タイムライン上に引く
  If line = 6 Then
    Set rngBase = Range(baseColumn & 5)
    
    '直線コネクタ生成
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "イナズマ線B_5"
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
  End If
  
  Set rngBase = Range(baseColumn & line)
  
  
  
  
  'イナズマ線を引かない場合は、基準日のみ引く
  If setVal("setLightning") = False Or Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_LateOrEarly") & line) = 0 Then
    
    '直線コネクタ生成
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "イナズマ線B_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
    Exit Function
  
  '進捗が0%以上の場合は、イナズマ線を引く
  ElseIf Range(setVal("cell_Progress") & line) >= 0 Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "イナズマ線S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("実績_" & line), 4
  
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.Top, rngBase.Left + 10, rngBase.Top + rngBase.Height).Select
    With Selection
      .Name = "イナズマ線S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes("実績_" & line), 4
    
    
'
'      startTask = "タスク_" & tmpLine
'      thisTask = "タスク_" & line
'
'    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
'    Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
'
'    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
'    Selection.Name = "イナズマ線_" & line
  
      
  End If
Exit Function





  If Range(setVal("GUNT_START_DAY") & line) <> "" Then
    startColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_START_DAY") & line))
  Else
    startColumn = baseColumn
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) <> "" Then
    endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
  
  ElseIf Range(setVal("GUNT_END_DAY") & line) <> "" Then
    endColumn = WBS_Option.日付セル検索(Range(setVal("GUNT_END_DAY") & line))
  Else
    endColumn = baseColumn
  End If
    
  'Shapeを配置するための基準となるセル
  Set rngStart = Range(startColumn & line)
  Set rngEnd = Range(endColumn & line)

  
  '遅早工数の値
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
    chkDay = WBS_Option.イナズマ線用日付計算(setVal("baseDay"), Range(setVal("cell_LateOrEarly") & line))
    chkDayColumn = WBS_Option.日付セル検索(chkDay)
    
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
' * タイムラインに追加
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タイムラインに追加(line As Long, Optional autoFlg As Boolean = False)
  Dim endLine As Long, colLine As Long, endColLine As Long
  Dim ShapeTopStart As Long, count As Long
  Dim shp As Shape
  Dim rng As Range
  Dim colorVal As Long
  Dim targetTaskName As String
  
  Const funcName As String = "Ctl_Chart.タイムラインに追加"


'  On Error GoTo catchError
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "function1")
  
  startColumn = WBS_Option.日付セル検索(Format(Range("M" & line), "yyyy/mm/dd"))
  endColumn = WBS_Option.日付セル検索(Format(Range("N" & line), "yyyy/mm/dd"))


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



  'タイムライン行の幅を広げる
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
  
  If startColumn = endColumn And (Range("K" & line) = "公開" Or Range("L" & line) = "公開") Then
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
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
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
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
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

    
    '重複削除
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
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * センター
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function センター()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseDate As Date
  Dim baseColumn As String
  
  
'  On Error GoTo catchError

  If setVal("startDay") >= setVal("baseDay") - 10 Then
    baseDate = setVal("startDay")
  Else
    baseDate = setVal("baseDay") - 10
    
  End If
  
  baseColumn = WBS_Option.日付セル検索(baseDate)
  Application.Goto Reference:=Range(baseColumn & 6), Scroll:=True


  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function





'**************************************************************************************************
' * ガントチャート選択
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
' * Chartから予定日を操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function changeShapes()
  Dim rng As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
  
  Call Library.showDebugForm("Chartから予定日を操作", "処理開始")
  
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With

  With ActiveSheet.Shapes(changeShapesName)
    Set rng = Range(.TopLeftCell, .BottomRightCell)
  End With
  If rng.Address(False, False) Like "*:*" Then
    tmp = Split(rng.Address(False, False), ":")
    changeShapesName = Replace(changeShapesName, "タスク_", "")
    mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName) = Range(getColumnName(Range(tmp(0)).Column) & 4)
    mainSheet.Range(setVal("GUNT_END_DAY") & changeShapesName) = Range(getColumnName(Range(tmp(1)).Column) & 4)
  Else
    tmp = rng.Address(False, False)
    changeShapesName = Replace(changeShapesName, "タスク_", "")
    mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
    mainSheet.Range(setVal("GUNT_END_DAY") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
  End If
  
  '先行タスクの終了日+1を開始日に設定
  newStartDay = mainSheet.Range(setVal("GUNT_START_DAY") & changeShapesName)
  Call init.chkHollyday(newStartDay, HollydayName)
  Do While HollydayName <> ""
    newStartDay = newStartDay - 1
    Call init.chkHollyday(newStartDay, HollydayName)
  Loop
  Range(setVal("GUNT_START_DAY") & changeShapesName) = newStartDay
  
  '終了日を再設定
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
  
  ActiveSheet.Shapes("タスク_" & changeShapesName).Delete
  If Range(setVal("cell_Task") & CLng(changeShapesName)) <> "" Then
    ActiveSheet.Shapes("先行タスク設定_" & CLng(changeShapesName)).Delete
  End If
  If setVal("viewGant_Assignor") = True Then
    ActiveSheet.Shapes("担当者_" & changeShapesName).Delete
  End If
  
  Call 計画線設定(CLng(changeShapesName))
  Call タスクのリンク設定(CLng(changeShapesName))
  
  changeShapesName = ""

  Call Library.showDebugForm("Chartから予定日を操作", "処理終了")
End Function


'**************************************************************************************************
' * タスクのリンク設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスクのリンク設定(line As Long)
  Dim startTask As String, thisTask As String
  Dim interval As Long
  Dim tmpLine As Variant
  
'  On Error GoTo catchError


  If Range(setVal("cell_Task") & line) = "" Then
    Exit Function
  Else
    For Each tmpLine In Split(Range(setVal("cell_Task") & line), ",")
      startTask = "タスク_" & tmpLine
      thisTask = "タスク_" & line
    
      'カギ線コネクタ生成
      ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
      Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
      Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startTask), 4
      Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
      Selection.Name = "先行タスク設定_" & line
      
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
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
Function 機能追加_計画線()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim shp As Shape
  
  
  
  Const funcName As String = "Ctl_Chart.機能追加_計画線"

  '処理開始--------------------------------------
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
        
        'B列と同じ色なら計画線と認識して処理
        If shp.line.ForeColor = Range("B" & line).Interior.Color Or shp.line.ForeColor = 65370 Then
          shp.Name = "taskLine_" & line
'          shp.OnAction = "beforeChangeShapes"
        End If
        
        Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "")
        
    End Select
    
  Next
  

  '処理終了--------------------------------------
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

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function







