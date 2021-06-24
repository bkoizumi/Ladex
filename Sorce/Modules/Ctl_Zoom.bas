Attribute VB_Name = "Ctl_Zoom"

'**************************************************************************************************
' * 選択セルの拡大表示/終了
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
    With Frm_Zoom
      .StartUpPosition = 0
      If topPosition = "" Then
        .Top = 10
        .Left = 120
      Else
        .Top = topPosition
        .Left = leftPosition
      End If
      .TextBox = cellVal
      .TextBox.MultiLine = True
      .TextBox.MultiLine = True
      .TextBox.EnterKeyBehavior = True
      .Label1.Caption = "選択セル：" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  
      .Show vbModeless
    End With
  
  Else
    If (Frm_Zoom.Visible = True) Then
      Frm_Zoom.TextBox.SelText = slctCellAddress
    End If
  End If
  

End Function


'==================================================================================================
Function ZoomOut(Text As String, SetTargetAddress As String)
  
  SetTargetAddress = Replace(SetTargetAddress, "選択セル：", "")
  Range(SetTargetAddress).Value = Text
  Call endScript
End Function

