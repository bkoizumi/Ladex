Attribute VB_Name = "Ctl_Stamp"

'==================================================================================================
Function ����_�ψ�()

  Dim ShapeName As String
  Dim ActvSheet As Worksheet
  Dim ActvCell As Range
  Dim objShp As Shape

  On Error GoTo catchError

  Call Library.startScript
  Call init.setting
  Set ActvSheet = ActiveSheet
  Set ActvCell = ActiveCell
  
  BK_sheetStamp.Activate
'  sheetsetting.Shapes.Range(Array("TB_���t")).Select
'  'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(Date, "yyyy/mm/dd")
'  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(Now(), "mm/dd hh:nn")
'
'  sheetsetting.Shapes.Range(Array("TB_���O")).Select
'  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = setVal("name")
'
'  sheetsetting.Shapes.Range(Array("TB_��")).Select
'  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = checkFlg
  
  BK_sheetStamp.Shapes.Range(Array("�ψ�")).Select
  Selection.Copy
  
  'Call Library.waitTime(1000)
'  ActiveWorkbook.Activate
  ActvSheet.Activate
  
  ActiveSheet.Range(ActvCell.Address).Select
  
  
  
  addShapeLeft = 0
  ShapeName = "�ψ�"
  
  
  
  ActiveSheet.PasteSpecial Format:="�} (PNG)", Link:=False, DisplayAsIcon:=False
  Selection.Placement = xlMoveAndSize
  Selection.ShapeRange.LockAspectRatio = msoFalse
  
  Selection.ShapeRange.Width = 15
  Selection.ShapeRange.Height = 15
  Selection.ShapeRange.Name = ShapeName
'  Selection.ShapeRange.IncrementLeft 2.8 + addShapeLeft
'  Selection.ShapeRange.IncrementTop 2.8
    
  ActiveSheet.Range(ActvCell.Address).Select
  
  Call Library.endScript
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, Err.Description, True)
  Call Library.endScript
End Function


'==================================================================================================
Function ����_�m�F��(Optional nameVal As String, Optional FontName As String, Optional ShapeName As String)

  
  Dim ActvSheet As Worksheet
  Dim ActvCell As Range
  Dim objShp As Shape

  On Error GoTo catchError

  Call Library.startScript
  Call init.setting
  Set ActvSheet = ActiveSheet
  Set ActvCell = ActiveCell
  

  
  
  BK_sheetStamp.Activate
'  sheetsetting.Shapes.Range(Array("TB_���t")).Select
'  'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(Date, "yyyy/mm/dd")
'  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(Now(), "mm/dd hh:nn")
'
  BK_sheetStamp.Shapes.Range(Array("TB_���O2")).Select
  If nameVal = "" Then
    nameVal = Library.getRegistry("Main", "StampVal")
  End If
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = nameVal
  
  BK_sheetStamp.Shapes.Range(Array("�F��")).Select
  Selection.Copy
  
  ActvSheet.Activate
  ActiveSheet.Range(ActvCell.Address).Select
  addShapeLeft = 0
  If ShapeName = "" Then
    ShapeName = "����_�m�F��"
  End If
  
  
  ActiveSheet.PasteSpecial Format:="�} (PNG)", Link:=False, DisplayAsIcon:=False
  Selection.Placement = xlMoveAndSize
  Selection.ShapeRange.LockAspectRatio = msoFalse
  
  Selection.ShapeRange.Width = 45
  Selection.ShapeRange.Height = 45
  Selection.ShapeRange.Name = ShapeName
'  Selection.ShapeRange.IncrementLeft 2.8 + addShapeLeft
'  Selection.ShapeRange.IncrementTop 2.8
    
  ActiveSheet.Range(ActvCell.Address).Select
  
  Call Library.endScript
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, Err.Description, True)
  Call Library.endScript
End Function
