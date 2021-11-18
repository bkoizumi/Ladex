Attribute VB_Name = "Ctl_Zoom"
Option Explicit

'**************************************************************************************************
' * �I���Z���̊g��\��/�I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ZoomIn(Optional slctCellAddress As String)
  Dim cellVal As String
  Dim topPosition As Long, leftPosition As Long
  Dim cellWidth As Long
  Dim targetBook As Workbook
  Dim targetSheet As Worksheet
  
  If slctCellAddress = "" Then
  End If
  
  If ActiveCell.HasFormula = False Then
    cellVal = ActiveCell.Text
  Else
    cellVal = ActiveCell.Formula
  End If
  Set targetBook = ActiveWorkbook
  Set targetSheet = ActiveSheet
  
  cellWidth = ActiveCell.Width
  If cellWidth <= 330 Then
    cellWidth = 330
  ElseIf cellWidth >= 400 And cellWidth > Application.Width Then
    cellWidth = 400
  End If
  
  With Frm_Zoom
    .Width = cellWidth + 40
    .TextBox.Width = cellWidth
    .TextBox = cellVal
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    
    If cellVal = StrConv(cellVal, vbNarrow) Then
      '���p�̏ꍇ
      .TextBox.IMEMode = fmIMEModeOff
    Else
      '�S�p�̏ꍇ
      .TextBox.IMEMode = fmIMEModeOn
    End If
    
    .Label1.Caption = "�I���Z���F" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    .Show vbModeless
  
  End With
End Function


'==================================================================================================
Function ZoomOut(Text As String, SetTargetAddress As String)
  
  SetTargetAddress = Replace(SetTargetAddress, "�I���Z���F", "")
  
  targetBook.Activate
  targetSheet.Activate
  
  Range(SetTargetAddress).Value = Text
  Call endScript
End Function


'==================================================================================================
'�S��ʕ\��
Function Zoom01()
  Dim topPosition As Long, leftPosition As Long
  
  Application.DisplayFullScreen = True
  
  topPosition = Library.getRegistry("UserForm", "Zoom01Top")
  leftPosition = Library.getRegistry("UserForm", "Zoom01Left")

  Call Ctl_UsrForm.�\���ʒu(topPosition, leftPosition)
  With Frm_DispFullScreenForm
    .StartUpPosition = 0
    .Top = topPosition
    .Left = leftPosition
    .Show vbModeless
  End With
  
End Function

