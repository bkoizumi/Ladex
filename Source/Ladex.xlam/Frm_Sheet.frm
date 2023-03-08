VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Sheet 
   Caption         =   "シート管理"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675.001
   OleObjectBlob   =   "Frm_Sheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InitializeFlg  As Boolean
Public selectLine     As Long



'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  '処理開始--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True
  
  With Frm_Sheet
    .Caption = "[" & thisAppName & "] シート管理 "
    .inputSheetName.Value = ActiveSheet.Name
    With SheetList
      .View = lvwReport
      .LabelEdit = lvwManual
      .HideSelection = False
      .AllowColumnReorder = True
      .FullRowSelect = True
      .Gridlines = True
      .ColumnHeaders.add , "_ID", "#", 30
      .ColumnHeaders.add , "_Display", "表示", 30, lvwColumnCenter
      .ColumnHeaders.add , "_SheetName", "シート名", 140
      
      For line = 1 To ActiveWorkbook.Worksheets.count
        With .ListItems.add
          .Text = line
          
          If ActiveWorkbook.Worksheets(line).Visible = True Then
            .SubItems(1) = "○"
          
          Else
            .SubItems(1) = "－"
          End If
          .SubItems(2) = ActiveWorkbook.Worksheets(line).Name
        End With
        
        If ActiveWorkbook.Worksheets(line).Name = ActiveSheet.Name Then
          selectLine = line
        End If
        DoEvents
      Next
      
      .ListItems(selectLine).EnsureVisible
      .ListItems(selectLine).Selected = True
      .SetFocus
    End With
  End With
  
'  add.Accelerator = "a"     'シート追加
'  edit.Accelerator = "r"    'シート名変更
'  display.Accelerator = "s" 'シート表示/非表示
'  del.Accelerator = "d"     'シート削除
'  copy.Accelerator = "c"     'シートコピー
  
    
  InitializeFlg = False
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub



'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub SheetList_Click()
  Dim sheetName As String, meg As String
  Const funcName As String = "Frm_Sheet.edit_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  sheetName = SheetList.SelectedItem.SubItems(2)
  Frm_Sheet.inputSheetName.Value = sheetName
  
  If ActiveWorkbook.Worksheets(sheetName).Visible = 2 Then
    meg = "マクロによって非表示となっているシートです" & vbNewLine & "マクロの動作に影響を与える可能性があります。"
    Frm_Sheet.add.Enabled = False
    Frm_Sheet.edit.Enabled = False
    Frm_Sheet.del.Enabled = False
    
    Frm_Sheet.up.Enabled = False
    Frm_Sheet.down.Enabled = False
    
  ElseIf ActiveWorkbook.Worksheets(sheetName).Visible = True Then
'    meg = "ダブルクリックで選択(アクティブ化)します"
    Frm_Sheet.add.Enabled = True
    Frm_Sheet.edit.Enabled = True
    Frm_Sheet.del.Enabled = True
  
    Frm_Sheet.up.Enabled = True
    Frm_Sheet.down.Enabled = True
    ActiveWorkbook.Worksheets(sheetName).Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
  Else
    meg = "非表示となっているシートです"
'    meg = meg & vbNewLine & "ダブルクリックで表示し、選択(アクティブ化)します"
    Frm_Sheet.add.Enabled = False
    Frm_Sheet.edit.Enabled = False
    Frm_Sheet.del.Enabled = True
    
    Frm_Sheet.up.Enabled = False
    Frm_Sheet.down.Enabled = False
    
  End If

    sheetInfo.Caption = meg & vbNewLine & ""
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
Private Sub SheetList_DblClick()
  Dim sheetName As String, sheetDspFLg As String
  
  Call Library.startScript
  selectLine = SheetList.SelectedItem.Text
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  sheetDspFLg = SheetList.SelectedItem.SubItems(1)
  
  If sheetDspFLg = "○" Then
    ActiveWorkbook.Worksheets(sheetName).Select
  Else
    ActiveWorkbook.Sheets(sheetName).Visible = True
    ActiveWorkbook.Worksheets(sheetName).Select
  End If
  
  Unload Me
  Call Library.endScript
End Sub

'==================================================================================================
'上
Private Sub up_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.up_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  Sheets(sheetName).Move Before:=Sheets(SheetList.SelectedItem.Text - 1)
  
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub


'==================================================================================================
'下
Private Sub down_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.down_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  Sheets(sheetName).Move Before:=Sheets(SheetList.SelectedItem.Text + 2)
  
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'シート名変更
Private Sub edit_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.edit_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  If Library.chkSheetExists(inputSheetName.Value) = False Then
    ActiveWorkbook.Sheets(sheetName).Select
    ActiveWorkbook.Sheets(sheetName).Name = inputSheetName.Value
  Else
    sheetInfo.Caption = inputSheetName.Value & "は、すでに存在します"
  End If
  
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'シート追加
Private Sub add_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.add_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  ActiveWorkbook.Sheets(sheetName).Select
  
  If Library.chkSheetExists(inputSheetName.Value) = False Then
    Sheets.add After:=ActiveSheet
    ActiveSheet.Name = inputSheetName.Value
  Else
    sheetInfo.Caption = inputSheetName.Value & "はすでに存在します"
  End If
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'シートコピー
Private Sub copy_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.copy_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  ActiveWorkbook.Sheets(sheetName).Select
  
  If Library.chkSheetExists(inputSheetName.Value) = False Then
    Call Library.startScript
    ActiveWorkbook.Sheets(sheetName).copy After:=Sheets(Worksheets.count)
    ActiveSheet.Name = inputSheetName.Value
    
    ActiveSheet.Move After:=ActiveWorkbook.Sheets(sheetName)
    
    
    Call Library.endScript
  Else
    sheetInfo.Caption = inputSheetName.Value & "はすでに存在します"
  End If
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'シート削除
Private Sub del_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.del_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  
  If MsgBox(sheetName & "を削除します(元にもどせません)", vbYesNo + vbExclamation) = vbYes Then
    ActiveWorkbook.Sheets(sheetName).delete
  End If
  
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub


'==================================================================================================
'表示/非表示
Private Sub display_Click()
  Dim sheetName As String
   
  Const funcName As String = "Frm_Sheet.display_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  sheetName = SheetList.SelectedItem.SubItems(2)
  If ActiveWorkbook.Sheets(sheetName).Visible = True Then
    ActiveWorkbook.Sheets(sheetName).Visible = False
  Else
    ActiveWorkbook.Sheets(sheetName).Visible = True
  End If
  Call reLoadList
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'複数シート生成
Private Sub addSheets_Click()
  Unload Me
  Call Ctl_Book.連続シート追加
  Call Ctl_Sheet.シート管理_フォーム表示
End Sub



'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
Function reLoadList()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.reLoadList"

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Call Library.startScript
  SheetList.ListItems.Clear
  SheetList.ColumnHeaders.Clear

  With SheetList
    .View = lvwReport
    .LabelEdit = lvwManual
    .HideSelection = False
    .AllowColumnReorder = True
    .FullRowSelect = True
    .Gridlines = True
    .ColumnHeaders.add , "_ID", "#", 30
    .ColumnHeaders.add , "_Display", "表示", 30, lvwColumnCenter
    .ColumnHeaders.add , "_SheetName", "シート名", 140
    
    For line = 1 To ActiveWorkbook.Worksheets.count
      With .ListItems.add
        .Text = line
        
        If ActiveWorkbook.Worksheets(line).Visible = True Then
          .SubItems(1) = "○"
        
        Else
          .SubItems(1) = "－"
        End If
        .SubItems(2) = ActiveWorkbook.Worksheets(line).Name
      End With
      
      If ActiveWorkbook.Worksheets(line).Name = ActiveSheet.Name Then
        selectLine = line
      End If
      'DoEvents
    Next
    DoEvents
    
    .ListItems(selectLine).EnsureVisible
    .ListItems(selectLine).Selected = True
    .SetFocus
  End With

  Call Library.endScript

End Function

