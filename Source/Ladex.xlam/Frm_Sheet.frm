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
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  Call Library.delSheetData(BK_sheetSheetList)
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  Application.Cursor = xlDefault
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
      .ColumnHeaders.add , "_Display", "表示", 30
      .ColumnHeaders.add , "_SheetName", "シート名", 140
      
      For line = 1 To ActiveWorkbook.Worksheets.count
        With .ListItems.add
          .Text = line
          If ActiveWorkbook.Worksheets(line).Visible = True Then
            .SubItems(1) = "○"
          End If
          .SubItems(2) = ActiveWorkbook.Worksheets(line).Name
        End With
        
        If ActiveWorkbook.Worksheets(line).Name = ActiveSheet.Name Then
          selectLine = line
        End If
        
        BK_sheetSheetList.Range("A" & line) = SheetList.ListItems.Item(line).Text
        BK_sheetSheetList.Range("B" & line) = SheetList.ListItems.Item(line).SubItems(1)
        BK_sheetSheetList.Range("C" & line) = SheetList.ListItems.Item(line).SubItems(2)
        
        BK_sheetSheetList.Range("D" & line) = SheetList.ListItems.Item(line).Text
        BK_sheetSheetList.Range("E" & line) = SheetList.ListItems.Item(line).SubItems(1)
        BK_sheetSheetList.Range("F" & line) = SheetList.ListItems.Item(line).SubItems(2)
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
  Call Library.endScript
  
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
  
  BK_sheetSheetList.Range("D" & selectLine) = BK_sheetSheetList.Range("D" & selectLine) - 1
  BK_sheetSheetList.Range("D" & selectLine - 1) = BK_sheetSheetList.Range("D" & selectLine - 1) + 1
  
  selectLine = selectLine - 1
  Call reLoadList
End Sub

'==================================================================================================
'下
Private Sub down_Click()
  selectLine = SheetList.SelectedItem.Text
  
  BK_sheetSheetList.Range("D" & selectLine) = BK_sheetSheetList.Range("D" & selectLine) + 1
  BK_sheetSheetList.Range("D" & selectLine + 1) = BK_sheetSheetList.Range("D" & selectLine + 1) - 1
  
  selectLine = selectLine + 1
  Call reLoadList
End Sub

'==================================================================================================
'シート名変更
Private Sub edit_Click()
  selectLine = SheetList.SelectedItem.Text
  
  BK_sheetSheetList.Range("F" & selectLine) = SheetName.Value
  
  Call reLoadList
End Sub

'==================================================================================================
'シート追加
Private Sub add_Click()
  Dim endLine As Long
  
  endLine = BK_sheetSheetList.Cells(Rows.count, 4).End(xlUp).Row + 1
  
  BK_sheetSheetList.Range("D" & endLine) = endLine
  BK_sheetSheetList.Range("E" & endLine) = "○"
  BK_sheetSheetList.Range("F" & endLine) = SheetName.Value
  
  selectLine = endLine
  Call reLoadList
End Sub

'==================================================================================================
'シート削除
Private Sub del_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "削除" Then
    If BK_sheetSheetList.Range("E" & selectLine) <> "" Then
      BK_sheetSheetList.Range("E" & selectLine) = BK_sheetSheetList.Range("B" & selectLine)
    Else
      BK_sheetSheetList.Range("E" & selectLine) = "○"
    End If
  Else
    BK_sheetSheetList.Range("E" & selectLine) = "削除"
  End If
  Call reLoadList
End Sub


'==================================================================================================
'表示/非表示
Private Sub display_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "○" Then
    BK_sheetSheetList.Range("E" & selectLine) = "X"
  Else
    BK_sheetSheetList.Range("E" & selectLine) = "○"
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
  
  SheetName = BK_sheetSheetList.Range("F" & selectLine).Value
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

  endLine = BK_sheetSheetList.Cells(Rows.count, 4).End(xlUp).Row
  Call Library.startScript
  
  For line = 1 To endLine
    If BK_sheetSheetList.Range("A" & line) <> BK_sheetSheetList.Range("D" & line) Then
      If BK_sheetSheetList.Range("A" & line) = "" And BK_sheetSheetList.Range("E" & line) <> "削除" Then
        '新規シート追加
        targetBook.Sheets.add After:=ActiveSheet
        targetBook.ActiveSheet.Name = BK_sheetSheetList.Range("F" & line).Value
        targetBook.ActiveSheet.Move After:=Sheets(BK_sheetSheetList.Range("D" & line - 1))
        
      Else
        'シートの順番変更
        targetBook.Sheets(BK_sheetSheetList.Range("A" & line)).Move before:=Sheets(BK_sheetSheetList.Range("D" & line))
      End If
    ElseIf BK_sheetSheetList.Range("B" & line) <> BK_sheetSheetList.Range("E" & line) Then
      'シートの表示/非表示切り替え
      If BK_sheetSheetList.Range("E" & line) = "○" Then
        targetBook.Sheets(BK_sheetSheetList.Range("F" & line).Value).Visible = True
      
      ElseIf BK_sheetSheetList.Range("E" & line) = "削除" Then
        targetBook.Worksheets(BK_sheetSheetList.Range("F" & line).Value).Select
        ActiveWindow.SelectedSheets.delete
      Else
        targetBook.Sheets(BK_sheetSheetList.Range("F" & line).Value).Visible = False
      End If
      
    ElseIf BK_sheetSheetList.Range("C" & line) <> BK_sheetSheetList.Range("F" & line) Then
      'シート名の変更
      targetBook.Sheets(BK_sheetSheetList.Range("C" & line).Value).Name = BK_sheetSheetList.Range("F" & line).Value
    End If
      
  Next
  
  Call Library.delSheetData(BK_sheetSheetList)
  Set targetBook = Nothing
  
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "○" And BK_sheetSheetList.Range("F" & selectLine).Value <> "" Then
    ActiveWorkbook.Worksheets(BK_sheetSheetList.Range("F" & selectLine).Value).Select
  End If
  
  
  Call Library.endScript
'  Unload Me
End Sub


'==================================================================================================
Function reLoadList()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  endLine = BK_sheetSheetList.Cells(Rows.count, 4).End(xlUp).Row

  BK_sheetSheetList.Sort.SortFields.Clear
  BK_sheetSheetList.Sort.SortFields.add Key:=Range("D1:D" & endLine), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With BK_sheetSheetList.Sort
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
      .ColumnHeaders.add , "_Display", "表示", 30
      .ColumnHeaders.add , "_SheetName", "シート名", 140
    
    For line = 1 To endLine
      With .ListItems.add
        .Text = BK_sheetSheetList.Range("D" & line).Value
        .SubItems(1) = BK_sheetSheetList.Range("E" & line).Value
        .SubItems(2) = BK_sheetSheetList.Range("F" & line).Value
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

