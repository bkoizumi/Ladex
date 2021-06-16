Attribute VB_Name = "Ctl_HighLight"
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

#End If

Public Type POINTAPI
    X As Long
    Y As Long
End Type


'==================================================================================================
Sub GetCursorposition()

  Dim p        As POINTAPI 'API用変数
  Dim Rng  As Range
  
  If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = False
  End If
  If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = False
  End If
  Call Library.waitTime(50)
  Call Library.startScript
  
  'カーソル位置取得
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.Y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.Y).Select
  End If
  
  If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = True
  End If
  If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = True
  End If
  
  Call Library.endScript


End Sub

'==================================================================================================
Function showStart(ByVal Target As Range, Optional targetArea_X As Range, Optional targetArea_Y As Range)
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  
  If Library.chkShapeName("HighLight_X") = True Then
    ActiveSheet.Shapes("HighLight_X").delete
  End If
  If Library.chkShapeName("HighLight_Y") = True Then
    ActiveSheet.Shapes("HighLight_Y").delete
  End If
  
  If BKh_rbPressed = True Then
    Set Rng = Range("A" & Target.Row)
    
    If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[X,B]" Then
      If targetArea_X Is Nothing Then
        Call showStart_X(Target, targetArea_X)
        
      ElseIf Not (targetArea_X Is Nothing) And Application.Intersect(ActiveCell, targetArea_X) Is Nothing Then
        Call showStart_X(Target, targetArea_X)
      End If
    End If
    
    If Library.getRegistry(RegistrySubKey, "HighLight_DspDirection") Like "[Y,B]" Then
      If targetArea_Y Is Nothing Then
        Call showStart_Y(Target, targetArea_Y)
      ElseIf Not (targetArea_Y Is Nothing) And Application.Intersect(ActiveCell, targetArea_Y) Is Nothing Then
        Call showStart_Y(Target, targetArea_Y)
      End If
    End If
  End If

  Target.Activate
  Set Rng = Nothing

End Function


'==================================================================================================
Function showStart_X(ByVal Target As Range, Optional targetArea_X As Range)
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  If BKh_rbPressed = True Then
    Set Rng = Range("A" & Target.Row)
    
    'MaxWidth = Application.Width
    'MaxWidth = Range(Cells(1, 1), Cells(1, Columns.count)).Width
    MaxWidth = 169056
     
    'Cells(Rows.count, Columns.Count)
    
    Set HighLight_X = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, _
      Left:=Rng.Left, Top:=Rng.Top, Width:=MaxWidth, Height:=Rng.Height)
    HighLight_X.Name = "HighLight_X"
    HighLight_X.OnAction = "GetCursorposition"
    
    '表示方法
    HighLight_DspMethod = Library.getRegistry("Main", "HighLight_DspMethod")
    
    '帯(塗りつぶし)
    If HighLight_DspMethod = "0" Then
      HighLight_X.Fill.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
'      Selection.ShapeRange.Width = Range(Cells(1, 1), Cells(1, Columns.count)).Width
    
    
    '囲み線
    ElseIf HighLight_DspMethod = "1" Then
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      Selection.ShapeRange.line.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
      Selection.ShapeRange.line.Weight = 3
    
    '直線
    ElseIf HighLight_DspMethod = "2" Then
'      Set Rng = Range("A" & Target.Row + 1)
      
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      Selection.ShapeRange.line.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
      Selection.ShapeRange.line.Weight = 3
      Selection.ShapeRange.Height = 1
'      Selection.ShapeRange.Top = Rng.Top
    
    End If
    
  
  End If
  Set Rng = Nothing

End Function


'==================================================================================================
Function showStart_Y(ByVal Target As Range, Optional targetArea_Y As Range)
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  If BKh_rbPressed = True Then
    MaxHeight = 169056
    
    
    Set Rng = Cells(1, Target.Column)
    
    Set HighLight_Y = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, _
      Left:=Rng.Left, Top:=Rng.Top, Width:=Rng.Width, Height:=MaxHeight)
      
    HighLight_Y.Name = "HighLight_Y"
    HighLight_Y.OnAction = "GetCursorposition"
    HighLight_Y.Fill.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
    
    
    '表示方法
    HighLight_DspMethod = Library.getRegistry("Main", "HighLight_DspMethod")
    
    '帯(塗りつぶし)
    If HighLight_DspMethod = "0" Then
      HighLight_Y.Fill.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
    
    '囲み線
    ElseIf HighLight_DspMethod = "1" Then
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      Selection.ShapeRange.line.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
      Selection.ShapeRange.line.Weight = 3
    
    '直線
    ElseIf HighLight_DspMethod = "2" Then
'      Set Rng = Cells(1, Target.Column + 1)
      
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = Library.getRegistry(RegistrySubKey, "HighLight_Color")
      Selection.ShapeRange.line.Transparency = Library.getRegistry("Main", "HighLight_TransparentRate") / 100
      Selection.ShapeRange.line.Weight = 3
      Selection.ShapeRange.Width = 1
'      Selection.ShapeRange.Left = Rng.Left
    
    End If
  
  End If
  Set Rng = Nothing

End Function



Sub Sample()
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    
    On Error Resume Next
    
    
    'Shapeを配置するための基準となるセル
    Set rngStart = Range("A2")
    Set rngEnd = Range("EEL2")
    
    'セルのLeft、Top、Widthプロパティを利用して位置決め
    BX = rngStart.Left
    BY = rngStart.Top
    EX = rngEnd.Left + rngEnd.Width
    EY = rngEnd.Top
    
    '直線
    ActiveSheet.Shapes.AddLine BX, BY, EX, EY

    rngEnd.Select
End Sub
