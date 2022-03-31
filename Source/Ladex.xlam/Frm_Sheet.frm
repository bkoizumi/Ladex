VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Sheet 
   Caption         =   "シート管理"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945.001
   OleObjectBlob   =   "Frm_Sheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InitializeFlg   As Boolean
Public selectLine   As Long

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
  Call Library.delSheetData(LadexSh_SheetList)
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True
  
  With Frm_Sheet
    .Caption = "シート管理 " & thisAppName
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
          LadexSh_SheetList.Range("A" & line) = line
          LadexSh_SheetList.Range("D" & line) = line
          
          If ActiveWorkbook.Worksheets(line).Visible = True Then
            .SubItems(1) = "○"
            LadexSh_SheetList.Range("B" & line) = "○"
            LadexSh_SheetList.Range("E" & line) = "○"
          End If
          .SubItems(2) = ActiveWorkbook.Worksheets(line).Name
          LadexSh_SheetList.Range("C" & line) = ActiveWorkbook.Worksheets(line).Name
          LadexSh_SheetList.Range("F" & line) = ActiveWorkbook.Worksheets(line).Name
        End With
        
        If ActiveWorkbook.Worksheets(line).Name = ActiveSheet.Name Then
          selectLine = line
        End If
        DoEvents
      Next
      
      '最終行に空白追加
        With .ListItems.add
          .Text = line
          .SubItems(1) = ""
          .SubItems(2) = "シート末尾"
        End With
      .ListItems(selectLine).EnsureVisible
      .ListItems(selectLine).Selected = True
      .SetFocus
    End With
  End With
  
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
  selectLine = SheetList.SelectedItem.Text
  
  If SheetList.SelectedItem.SubItems(2) = "シート末尾" Then
    SheetName.Value = ""
    Exit Sub
  Else
    SheetName.Value = SheetList.SelectedItem.SubItems(2)
  End If
  
  If ActiveWorkbook.Worksheets(selectLine).Visible = 2 Then
    sheetInfo.Caption = "マクロによって非表示となっているシートです" & vbNewLine & "マクロの動作に影響を与える可能性があります。"
  ElseIf ActiveWorkbook.Worksheets(selectLine).Visible = True Then
    sheetInfo.Caption = "ダブルクリックで選択(アクティブ化)します"
  Else
    sheetInfo.Caption = ""
  End If
End Sub

'==================================================================================================
Private Sub SheetList_DblClick()
  Dim SheetName As String, sheetDspFLg As String
  
  Call Library.startScript
  selectLine = SheetList.SelectedItem.Text
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  sheetDspFLg = SheetList.SelectedItem.SubItems(1)
  
  If sheetDspFLg = "○" Then
    ActiveWorkbook.Worksheets(SheetName).Select
  Else
    targetBook.Sheets(SheetName).Visible = True
    ActiveWorkbook.Worksheets(SheetName).Select
  End If
  
  Unload Me
  Call Library.endScript
End Sub

'==================================================================================================
'上
Private Sub up_Click()
  selectLine = SheetList.SelectedItem.Text
  
  LadexSh_SheetList.Range("D" & selectLine) = LadexSh_SheetList.Range("D" & selectLine) - 1
  LadexSh_SheetList.Range("D" & selectLine - 1) = LadexSh_SheetList.Range("D" & selectLine - 1) + 1
  
  selectLine = selectLine - 1
  Call reLoadList
End Sub

'==================================================================================================
'下
Private Sub down_Click()
  selectLine = SheetList.SelectedItem.Text
  
  LadexSh_SheetList.Range("D" & selectLine) = LadexSh_SheetList.Range("D" & selectLine) + 1
  LadexSh_SheetList.Range("D" & selectLine + 1) = LadexSh_SheetList.Range("D" & selectLine + 1) - 1
  
  selectLine = selectLine + 1
  Call reLoadList
End Sub

'==================================================================================================
'シート名変更
Private Sub edit_Click()
  selectLine = SheetList.SelectedItem.Text
  
  LadexSh_SheetList.Range("F" & selectLine) = SheetName.Value
  
  Call reLoadList
End Sub

'==================================================================================================
'シート追加
Private Sub add_Click()
  Dim endLine As Long
  
  endLine = LadexSh_SheetList.Cells(Rows.count, 4).End(xlUp).Row + 1
  
  LadexSh_SheetList.Range("D" & endLine) = endLine
  LadexSh_SheetList.Range("E" & endLine) = "○"
  LadexSh_SheetList.Range("F" & endLine) = SheetName.Value
  
  selectLine = endLine
  Call reLoadList
End Sub

'==================================================================================================
'シート削除
Private Sub del_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If LadexSh_SheetList.Range("E" & selectLine) = "削除" Then
    If LadexSh_SheetList.Range("E" & selectLine) <> "" Then
      LadexSh_SheetList.Range("E" & selectLine) = LadexSh_SheetList.Range("B" & selectLine)
    Else
      LadexSh_SheetList.Range("E" & selectLine) = "○"
    End If
  Else
    LadexSh_SheetList.Range("E" & selectLine) = "削除"
  End If
  Call reLoadList
End Sub


'==================================================================================================
'表示/非表示
Private Sub display_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If LadexSh_SheetList.Range("E" & selectLine) = "○" Then
    LadexSh_SheetList.Range("E" & selectLine) = "X"
  Else
    LadexSh_SheetList.Range("E" & selectLine) = "○"
  End If
  Call reLoadList
End Sub

'==================================================================================================
'シートの選択
Private Sub active_Click()
  Dim SheetName As String, sheetDspFLg As String
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  sheetDspFLg = SheetList.SelectedItem.SubItems(1)
  selectLine = SheetList.SelectedItem.Text
  
  SheetName = LadexSh_SheetList.Range("F" & selectLine).Value
  If ActiveWorkbook.Worksheets(SheetName).Visible = True Then
    ActiveWorkbook.Worksheets(SheetName).Select
  Else
    ActiveWorkbook.Worksheets(SheetName).Visible = True
    ActiveWorkbook.Worksheets(SheetName).Select
  End If
  
  'Unload Me
End Sub

'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
' 実行
Private Sub Submit_Click()
  Dim line As Long, endLine As Long
  Dim selectLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  endLine = LadexSh_SheetList.Cells(Rows.count, 4).End(xlUp).Row
  Call Library.startScript
  
  For line = 1 To endLine
    If LadexSh_SheetList.Range("A" & line) <> LadexSh_SheetList.Range("D" & line) Then
      If LadexSh_SheetList.Range("A" & line) = "" And LadexSh_SheetList.Range("E" & line) <> "削除" Then
        '新規シート追加
        targetBook.Sheets.add After:=ActiveSheet
        targetBook.ActiveSheet.Name = LadexSh_SheetList.Range("F" & line).Value
        targetBook.ActiveSheet.Move After:=Sheets(LadexSh_SheetList.Range("D" & line - 1))
        
      Else
        'シートの順番変更
        targetBook.Sheets(LadexSh_SheetList.Range("A" & line)).Move before:=Sheets(LadexSh_SheetList.Range("D" & line))
      End If
    ElseIf LadexSh_SheetList.Range("B" & line) <> LadexSh_SheetList.Range("E" & line) Then
      'シートの表示/非表示切り替え
      If LadexSh_SheetList.Range("E" & line) = "○" Then
        targetBook.Sheets(LadexSh_SheetList.Range("F" & line).Value).Visible = True
      
      ElseIf LadexSh_SheetList.Range("E" & line) = "削除" Then
        targetBook.Worksheets(LadexSh_SheetList.Range("F" & line).Value).Select
        ActiveWindow.SelectedSheets.delete
      Else
        targetBook.Sheets(LadexSh_SheetList.Range("F" & line).Value).Visible = False
      End If
      
    ElseIf LadexSh_SheetList.Range("C" & line) <> LadexSh_SheetList.Range("F" & line) Then
      'シート名の変更
      targetBook.Sheets(LadexSh_SheetList.Range("C" & line).Value).Name = LadexSh_SheetList.Range("F" & line).Value
    End If
      
  Next
  
  Call Library.delSheetData(LadexSh_SheetList)
  Set targetBook = Nothing
  
  selectLine = SheetList.SelectedItem.Text
  
  If LadexSh_SheetList.Range("E" & selectLine) = "○" And LadexSh_SheetList.Range("F" & selectLine).Value <> "" Then
    ActiveWorkbook.Worksheets(LadexSh_SheetList.Range("F" & selectLine).Value).Select
  End If
  
  
  Call Library.endScript
'  Unload Me
End Sub


'==================================================================================================
Function reLoadList()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  endLine = LadexSh_SheetList.Cells(Rows.count, 4).End(xlUp).Row

  LadexSh_SheetList.Sort.SortFields.Clear
  LadexSh_SheetList.Sort.SortFields.add Key:=Range("D1:D" & endLine), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With LadexSh_SheetList.Sort
    .SetRange Range("A1:F" & endLine)
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
  
    
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
    
    For line = 1 To endLine
      With .ListItems.add
        .Text = LadexSh_SheetList.Range("D" & line).Value
        .SubItems(1) = LadexSh_SheetList.Range("E" & line).Value
        .SubItems(2) = LadexSh_SheetList.Range("F" & line).Value
      End With
    Next
    '最終行に空白追加
      With .ListItems.add
        .Text = line
        .SubItems(1) = ""
        .SubItems(2) = "シート末尾"
      End With
    
    .ListItems(selectLine).EnsureVisible
    .ListItems(selectLine).Selected = True
    .SetFocus
  End With
  
End Function

