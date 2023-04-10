Attribute VB_Name = "Ctl_HighLight"
Option Explicit

#If VBA7 And Win64 Then
  Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As LongPtr) As Long
#Else
  Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long
#End If

Dim HighLightColor As String
Dim HighLightDspDirection As String
Dim HighLightDspMethod As String
Dim HighlightTransparentRate As Long


Public Type POINTAPI
    X As Long
    y As Long
End Type


'==================================================================================================
Sub getCursorPosition()

  Dim p    As POINTAPI 'API用変数
  Dim Rng  As Range
  
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = False
  End If
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = False
  End If
  Call Library.waitTime(50)
  Call Library.startScript
  
  'カーソル位置取得
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.y).Select
  End If
  
'  Call Ctl_HighLight.showStart(ActiveCell, True)
  
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = True
  End If
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = True
  End If
  
  Call Library.endScript


End Sub

'==================================================================================================
Function showStart(ByVal Target As Range, Optional DelFlg As Boolean = True, Optional targetSheet As Worksheet)

  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  Call init.setting
  Call Library.startScript
  
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  
  If DelFlg = True Then
    If Library.chkShapeName("HighLight_X") = True Then
      targetSheet.Shapes("HighLight_X").delete
    End If
    If Library.chkShapeName("HighLight_Y") = True Then
      targetSheet.Shapes("HighLight_Y").delete
    End If
  End If
  
  If BKh_rbPressed = True Then
    Set Rng = Range("A" & Target.Row)
    
    HighLightColor = Library.getRegistry("Main", "HighLightColor")
    HighLightDspDirection = Library.getRegistry("Main", "HighLightDspDirection")
    HighLightDspMethod = Library.getRegistry("Main", "HighLightDspMethod")
    HighlightTransparentRate = CLng(Library.getRegistry("Main", "HighLightTransparentRate"))
    
    If ActiveWorkbook.Name Like "*WBS*" And Target.Column >= 23 Then
      Call showStart_X(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate, ActiveSheet)
      Call showStart_Y(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate, ActiveSheet)
    Else
      If HighLightDspDirection Like "[X,B]" Then
        Call showStart_X(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate, ActiveSheet)
      End If
      
      If HighLightDspDirection Like "[Y,B]" Then
        Call showStart_Y(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate, ActiveSheet)
      End If
    End If
  End If

  Target.Activate
  Set Rng = Nothing

  Call Library.endScript
End Function


'==================================================================================================
Function showStart_X(ByVal Target As Range, _
                  HighLightColor As String, _
                  HighLightDspDirection As String, _
                  HighLightDspMethod As String, _
                  HighlightTransparentRate As Long, _
                  Optional targetSheet As Worksheet)
                  
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  Dim MaxWidth As Long
  Dim HighLight_X
  
  If BKh_rbPressed = True Then
    If Target.Rows.count > 1 Then
      Set Rng = Range("A" & Target.Row & ":A" & Target.Row + Target.Rows.count - 1)
    Else
      Set Rng = Range("A" & Target.Row)
    End If
    
    'MaxWidth = Application.Width
    'MaxWidth = Range(Cells(1, 1), Cells(1, Columns.count)).Width
    MaxWidth = 169056
     
    'Cells(Rows.count, Columns.Count)
    
    Set HighLight_X = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, _
      Left:=Rng.Left, Top:=Rng.Top, Width:=MaxWidth, Height:=Rng.Height)
    HighLight_X.Name = "HighLight_X"
    HighLight_X.OnAction = "getCursorPosition"
    
    '表示方法
    HighLightDspMethod = HighLightDspMethod
    
    '帯(塗りつぶし)
    If HighLightDspMethod = "0" Then
      HighLight_X.Fill.ForeColor.RGB = HighLightColor
      targetSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Visible = msoTrue
      Selection.ShapeRange.Fill.Solid
      Selection.ShapeRange.Fill.Transparency = HighlightTransparentRate / 100
    
    
    '囲み線
    ElseIf HighLightDspMethod = "1" Then
      targetSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
    
    '直線
    ElseIf HighLightDspMethod = "2" Then
'      Set Rng = Range("A" & Target.Row + 1)
      
      targetSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
      Selection.ShapeRange.Height = 1
'      Selection.ShapeRange.Top = Rng.Top
    
    End If
    
  
  End If
  Set Rng = Nothing

End Function


'==================================================================================================
Function showStart_Y(ByVal Target As Range, _
                  HighLightColor As String, _
                  HighLightDspDirection As String, _
                  HighLightDspMethod As String, _
                  HighlightTransparentRate As Long, _
                  Optional targetSheet As Worksheet)
                   
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  Dim MaxHeight As Long
  Dim HighLight_Y
  
  If BKh_rbPressed = True Then
    MaxHeight = 169056
    
    If Target.Rows.count > 1 Then
      Set Rng = Range(Cells(1, Target.Column), Cells(1, Target.Column + Target.Columns.count - 1))
    Else
      Set Rng = Cells(1, Target.Column)
    End If
    
    
    Set HighLight_Y = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, _
      Left:=Rng.Left, Top:=Rng.Top, Width:=Rng.Width, Height:=MaxHeight)
      
    HighLight_Y.Name = "HighLight_Y"
    HighLight_Y.OnAction = "getCursorPosition"
    
    
    '表示方法--------------------------------------------------------------------------------------
    '帯(塗りつぶし)
    If HighLightDspMethod = "0" Then
      HighLight_Y.Fill.ForeColor.RGB = HighLightColor
      targetSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Transparency = HighlightTransparentRate / 100
    
    '囲み線
    ElseIf HighLightDspMethod = "1" Then
      targetSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
    
    '直線
    ElseIf HighLightDspMethod = "2" Then
      
      targetSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
      Selection.ShapeRange.Height = MaxHeight
      Selection.ShapeRange.Width = 0
'      Selection.ShapeRange.Top = Rng.Top
    
    End If
  
  End If
  Set Rng = Nothing

End Function



'==================================================================================================
Sub sample()
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, ex As Single, EY As Single
    
    On Error Resume Next
    
    
    'Shapeを配置するための基準となるセル
    Set rngStart = Range("A2")
    Set rngEnd = Range("EEL2")
    
    'セルのLeft、Top、Widthプロパティを利用して位置決め
    BX = rngStart.Left
    BY = rngStart.Top
    ex = rngEnd.Left + rngEnd.Width
    EY = rngEnd.Top
    
    '直線
    ActiveSheet.Shapes.AddLine BX, BY, ex, EY

    rngEnd.Select
End Sub

