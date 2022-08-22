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
    .TextBox.IMEMode = fmIMEModeOn

'    .TextBox.Font.Name = ActiveCell.Font.Name
'    .TextBox.Font.Name = LadexsetVal("BaseFont")
    .TextBox.Font.Name = "���C���I"
    
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
  Call Library.endScript
End Function


'==================================================================================================
'�S��ʕ\��
Function Zoom01()
  Dim topPosition As Long, leftPosition As Long
  
  Application.DisplayFullScreen = True
'  With Frm_DispFullScreenForm
'    .StartUpPosition = 3
'    .Show vbModeless
'  End With
  
End Function

'==================================================================================================
'�����Y�[���l�ɕύX
Function defaultZoom()
  Call init.setting
  Call Library.startScript
  
  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = Library.getRegistry("Main", "ZoomLevel")
  
  Call Library.endScript
End Function

