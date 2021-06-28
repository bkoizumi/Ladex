Attribute VB_Name = "Ctl_Zoom"

'**************************************************************************************************
' * �I���Z���̊g��\��/�I��
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function ZoomIn(Optional slctCellAddress As String)
  Dim cellVal As String
  
  If slctCellAddress = "" Then
    topPosition = Library.getRegistry("UserForm", "ZoomTop")
    leftPosition = Library.getRegistry("UserForm", "ZoomLeft")
    
    If ActiveCell.HasFormula = False Then
      cellVal = ActiveCell.Text
    Else
      cellVal = ActiveCell.Formula
    End If
    
'    Call Ctl_DefaultVal.setVal("ZoomIn", cellVal)
    
    With Frm_Zoom
      If topPosition = 0 Then
        .StartUpPosition = 2
      Else
        .StartUpPosition = 0
        .Top = topPosition
        .Left = leftPosition
      End If
      .TextBox = cellVal
      .TextBox.MultiLine = True
      .TextBox.MultiLine = True
      .TextBox.EnterKeyBehavior = True
      .Label1.Caption = "�I���Z���F" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  
      .Show vbModeless
    End With
  
  Else
    If (Frm_Zoom.Visible = True) Then
      'Frm_Zoom.TextBox.SelText = Ctl_DefaultVal.getVal("ZoomIn") & slctCellAddress
'      Frm_Zoom.TextBox = Ctl_DefaultVal.getVal("ZoomIn") & slctCellAddress
    End If
  End If
  

End Function


'==================================================================================================
Function ZoomOut(Text As String, SetTargetAddress As String)
  
  SetTargetAddress = Replace(SetTargetAddress, "�I���Z���F", "")
  
  Range(SetTargetAddress).Value = Text
  Call endScript
End Function


'==================================================================================================
'�S��ʕ\��

Function Zoom01()

  Call Library.startScript
  Application.DisplayFullScreen = True
  Call Library.endScript
  
  topPosition = Library.getRegistry("UserForm", "Zoom01Top")
  leftPosition = Library.getRegistry("UserForm", "Zoom01Left")
      
  With Frm_DispFullScreenForm
    
    If topPosition = 0 Then
      .StartUpPosition = 2
    Else
      .StartUpPosition = 0
      .Top = topPosition
      .Left = leftPosition
    End If
    .Show vbModeless
  End With
  
End Function

