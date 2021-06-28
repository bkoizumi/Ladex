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
Sub getCursorPosition()

  Dim p    As POINTAPI 'API�p�ϐ�
  Dim Rng  As Range
  
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = False
  End If
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = False
  End If
  Call Library.waitTime(50)
  Call Library.startScript
  
  '�J�[�\���ʒu�擾
  GetCursorPos p
  If TypeName(ActiveWindow.RangeFromPoint(p.X, p.Y)) = "Range" Then
    ActiveWindow.RangeFromPoint(p.X, p.Y).Select
  End If
  
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[X,B]" Then
    ActiveSheet.Shapes("HighLight_X").Visible = True
  End If
  If Library.getRegistry(RegistrySubKey, "HighLightDspDirection") Like "[Y,B]" Then
    ActiveSheet.Shapes("HighLight_Y").Visible = True
  End If
  
  Call Library.endScript


End Sub

'==================================================================================================
Function showStart(ByVal Target As Range, _
                  Optional HighLightColor As String, _
                  Optional HighLightDspDirection As String, _
                  Optional HighLightDspMethod As String, _
                  Optional HighlightTransparentRate As Long)
                  
                  
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  Call Library.startScript
  If Library.chkShapeName("HighLight_X") = True Then
    ActiveSheet.Shapes("HighLight_X").delete
  End If
  If Library.chkShapeName("HighLight_Y") = True Then
    ActiveSheet.Shapes("HighLight_Y").delete
  End If
  
  If BKh_rbPressed = True Then
    Set Rng = Range("A" & Target.Row)
    
        
    If HighLightColor = "" Then
      HighLightColor = Library.getRegistry("Main", "HighLightColor")
    End If
    
    If HighLightDspDirection = "" Then
      HighLightDspDirection = Library.getRegistry("Main", "HighLightDspDirection")
    End If
        
    If HighLightDspMethod = "" Then
      HighLightDspMethod = Library.getRegistry("Main", "HighLightDspMethod")
    End If
        
    If HighlightTransparentRate = 0 Then
      HighlightTransparentRate = CLng(Library.getRegistry("Main", "HighLightTransparentRate"))
    End If
    
    
    If HighLightDspDirection Like "[X,B]" Then
      Call showStart_X(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate)
    End If
    
    If HighLightDspDirection Like "[Y,B]" Then
      Call showStart_Y(Target, HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate)
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
                  HighlightTransparentRate As Long)
                  
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
    HighLight_X.OnAction = "getCursorPosition"
    
    '�\�����@
    HighLightDspMethod = HighLightDspMethod
    
    '��(�h��Ԃ�)
    If HighLightDspMethod = "0" Then
      HighLight_X.Fill.ForeColor.RGB = HighLightColor
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Transparency = HighlightTransparentRate / 100
'      Selection.ShapeRange.Width = Range(Cells(1, 1), Cells(1, Columns.count)).Width
    
    
    '�͂ݐ�
    ElseIf HighLightDspMethod = "1" Then
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
    
    '����
    ElseIf HighLightDspMethod = "2" Then
'      Set Rng = Range("A" & Target.Row + 1)
      
      ActiveSheet.Shapes.Range(Array("HighLight_X")).Select
      
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
                  HighlightTransparentRate As Long)
                   
  Dim Rng  As Range
  Dim ActvCellTop As Long, ActvCellLeft As Long
  
  If BKh_rbPressed = True Then
    MaxHeight = 169056
    
    
    Set Rng = Cells(1, Target.Column)
    
    Set HighLight_Y = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, _
      Left:=Rng.Left, Top:=Rng.Top, Width:=Rng.Width, Height:=MaxHeight)
      
    HighLight_Y.Name = "HighLight_Y"
    HighLight_Y.OnAction = "getCursorPosition"
    
    
    '�\�����@--------------------------------------------------------------------------------------
    '��(�h��Ԃ�)
    If HighLightDspMethod = "0" Then
      HighLight_Y.Fill.ForeColor.RGB = HighLightColor
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.line.Visible = msoFalse
      Selection.ShapeRange.Fill.Transparency = HighlightTransparentRate / 100
    
    '�͂ݐ�
    ElseIf HighLightDspMethod = "1" Then
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
      Selection.ShapeRange.Fill.Visible = msoFalse
      Selection.ShapeRange.line.Visible = msoTrue
      Selection.ShapeRange.line.ForeColor.RGB = HighLightColor
      Selection.ShapeRange.line.Transparency = HighlightTransparentRate / 100
      Selection.ShapeRange.line.Weight = 3
    
    '����
    ElseIf HighLightDspMethod = "2" Then
      
      ActiveSheet.Shapes.Range(Array("HighLight_Y")).Select
      
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



Sub Sample()
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    
    On Error Resume Next
    
    
    'Shape��z�u���邽�߂̊�ƂȂ�Z��
    Set rngStart = Range("A2")
    Set rngEnd = Range("EEL2")
    
    '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
    BX = rngStart.Left
    BY = rngStart.Top
    EX = rngEnd.Left + rngEnd.Width
    EY = rngEnd.Top
    
    '����
    ActiveSheet.Shapes.AddLine BX, BY, EX, EY

    rngEnd.Select
End Sub
