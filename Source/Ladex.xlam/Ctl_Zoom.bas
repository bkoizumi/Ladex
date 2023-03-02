Attribute VB_Name = "Ctl_Zoom"
Option Explicit

'**************************************************************************************************
' * 選択セルの拡大表示/終了
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
'    .TextBox.Font.Name = dicVal("BaseFont")
    .TextBox.Font.Name = "メイリオ"
    
    .Label1.Caption = "選択セル：" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    .Show vbModeless
  
  End With
End Function


'==================================================================================================
Function ZoomOut(Text As String, SetTargetAddress As String)
  On Error Resume Next
  
  SetTargetAddress = Replace(SetTargetAddress, "選択セル：", "")
  
  targetBook.Activate
  targetSheet.Activate
  
  Range(SetTargetAddress).Value = Text
  Call Library.endScript
End Function


'==================================================================================================
'全画面表示
Function 全画面表示()
  Dim topPosition As Long, leftPosition As Long
  
  Application.DisplayFullScreen = True
'  With Frm_DispFullScreenForm
'    .StartUpPosition = 3
'    .Show vbModeless
'  End With
  
End Function

'==================================================================================================
'初期ズーム値に変更
Function 初期表示倍率()
  Call init.setting
  Call Library.startScript
  
  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = Library.getRegistry("Main", "ZoomLevel")
  
  Call Library.endScript
End Function


'==================================================================================================
Function 指定倍率()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Zoom.指定倍率"

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

  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = Library.getRegistry("Main", "SpecifyZoomLevel")
  




  '処理終了--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
