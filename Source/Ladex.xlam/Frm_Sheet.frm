VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Sheet 
   Caption         =   "�V�[�g�Ǘ�"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945.001
   OleObjectBlob   =   "Frm_Sheet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  '�����J�n--------------------------------------
'  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  Call Library.delSheetData(BK_sheetSheetList)
  
  '�\���ʒu�w��----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  Application.Cursor = xlDefault
  InitializeFlg = True
  
  With Frm_Sheet
    .Caption = "�V�[�g�Ǘ� " & thisAppName
    With SheetList
      .View = lvwReport
      .LabelEdit = lvwManual
      .HideSelection = False
      .AllowColumnReorder = True
      .FullRowSelect = True
      .Gridlines = True
      .ColumnHeaders.add , "_ID", "#", 30
      .ColumnHeaders.add , "_Display", "�\��", 30
      .ColumnHeaders.add , "_SheetName", "�V�[�g��", 140
      
      For line = 1 To ActiveWorkbook.Worksheets.count
        With .ListItems.add
          .Text = line
          If ActiveWorkbook.Worksheets(line).Visible = True Then
            .SubItems(1) = "��"
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
      
      '�ŏI�s�ɋ󔒒ǉ�
        With .ListItems.add
          .Text = line
          .SubItems(1) = ""
          .SubItems(2) = "�V�[�g����"
        End With
      .ListItems(selectLine).EnsureVisible
      .ListItems(selectLine).Selected = True
      .SetFocus
    End With
  End With
  
  InitializeFlg = False
  Call Library.endScript
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub



'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub SheetList_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If SheetList.SelectedItem.SubItems(2) = "�V�[�g����" Then
    SheetName.Value = ""
    Exit Sub
  Else
    SheetName.Value = SheetList.SelectedItem.SubItems(2)
  End If
  
  If ActiveWorkbook.Worksheets(selectLine).Visible = 2 Then
    sheetInfo.Caption = "�}�N���ɂ���Ĕ�\���ƂȂ��Ă���V�[�g�ł�" & vbNewLine & "�}�N���̓���ɉe����^����\��������܂��B"
  ElseIf ActiveWorkbook.Worksheets(selectLine).Visible = True Then
    sheetInfo.Caption = "�_�u���N���b�N�őI��(�A�N�e�B�u��)���܂�"
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
  
  If sheetDspFLg = "��" Then
    ActiveWorkbook.Worksheets(SheetName).Select
  Else
    targetBook.Sheets(SheetName).Visible = True
    ActiveWorkbook.Worksheets(SheetName).Select
  End If
  
  Unload Me
  Call Library.endScript
End Sub

'==================================================================================================
'��
Private Sub up_Click()
  selectLine = SheetList.SelectedItem.Text
  
  BK_sheetSheetList.Range("D" & selectLine) = BK_sheetSheetList.Range("D" & selectLine) - 1
  BK_sheetSheetList.Range("D" & selectLine - 1) = BK_sheetSheetList.Range("D" & selectLine - 1) + 1
  
  selectLine = selectLine - 1
  Call reLoadList
End Sub

'==================================================================================================
'��
Private Sub down_Click()
  selectLine = SheetList.SelectedItem.Text
  
  BK_sheetSheetList.Range("D" & selectLine) = BK_sheetSheetList.Range("D" & selectLine) + 1
  BK_sheetSheetList.Range("D" & selectLine + 1) = BK_sheetSheetList.Range("D" & selectLine + 1) - 1
  
  selectLine = selectLine + 1
  Call reLoadList
End Sub

'==================================================================================================
'�V�[�g���ύX
Private Sub edit_Click()
  selectLine = SheetList.SelectedItem.Text
  
  BK_sheetSheetList.Range("F" & selectLine) = SheetName.Value
  
  Call reLoadList
End Sub

'==================================================================================================
'�V�[�g�ǉ�
Private Sub add_Click()
  Dim endLine As Long
  
  endLine = BK_sheetSheetList.Cells(Rows.count, 4).End(xlUp).Row + 1
  
  BK_sheetSheetList.Range("D" & endLine) = endLine
  BK_sheetSheetList.Range("E" & endLine) = "��"
  BK_sheetSheetList.Range("F" & endLine) = SheetName.Value
  
  selectLine = endLine
  Call reLoadList
End Sub

'==================================================================================================
'�V�[�g�폜
Private Sub del_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "�폜" Then
    If BK_sheetSheetList.Range("E" & selectLine) <> "" Then
      BK_sheetSheetList.Range("E" & selectLine) = BK_sheetSheetList.Range("B" & selectLine)
    Else
      BK_sheetSheetList.Range("E" & selectLine) = "��"
    End If
  Else
    BK_sheetSheetList.Range("E" & selectLine) = "�폜"
  End If
  Call reLoadList
End Sub


'==================================================================================================
'�\��/��\��
Private Sub display_Click()
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "��" Then
    BK_sheetSheetList.Range("E" & selectLine) = "X"
  Else
    BK_sheetSheetList.Range("E" & selectLine) = "��"
  End If
  Call reLoadList
End Sub

'==================================================================================================
'�V�[�g�̑I��
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
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
' ���s
Private Sub Submit_Click()
  Dim line As Long, endLine As Long
  Dim selectLine As Long
  Const funcName As String = "Frm_Sheet.UserForm_Initialize"

  endLine = BK_sheetSheetList.Cells(Rows.count, 4).End(xlUp).Row
  Call Library.startScript
  
  For line = 1 To endLine
    If BK_sheetSheetList.Range("A" & line) <> BK_sheetSheetList.Range("D" & line) Then
      If BK_sheetSheetList.Range("A" & line) = "" And BK_sheetSheetList.Range("E" & line) <> "�폜" Then
        '�V�K�V�[�g�ǉ�
        targetBook.Sheets.add After:=ActiveSheet
        targetBook.ActiveSheet.Name = BK_sheetSheetList.Range("F" & line).Value
        targetBook.ActiveSheet.Move After:=Sheets(BK_sheetSheetList.Range("D" & line - 1))
        
      Else
        '�V�[�g�̏��ԕύX
        targetBook.Sheets(BK_sheetSheetList.Range("A" & line)).Move before:=Sheets(BK_sheetSheetList.Range("D" & line))
      End If
    ElseIf BK_sheetSheetList.Range("B" & line) <> BK_sheetSheetList.Range("E" & line) Then
      '�V�[�g�̕\��/��\���؂�ւ�
      If BK_sheetSheetList.Range("E" & line) = "��" Then
        targetBook.Sheets(BK_sheetSheetList.Range("F" & line).Value).Visible = True
      
      ElseIf BK_sheetSheetList.Range("E" & line) = "�폜" Then
        targetBook.Worksheets(BK_sheetSheetList.Range("F" & line).Value).Select
        ActiveWindow.SelectedSheets.delete
      Else
        targetBook.Sheets(BK_sheetSheetList.Range("F" & line).Value).Visible = False
      End If
      
    ElseIf BK_sheetSheetList.Range("C" & line) <> BK_sheetSheetList.Range("F" & line) Then
      '�V�[�g���̕ύX
      targetBook.Sheets(BK_sheetSheetList.Range("C" & line).Value).Name = BK_sheetSheetList.Range("F" & line).Value
    End If
      
  Next
  
  Call Library.delSheetData(BK_sheetSheetList)
  Set targetBook = Nothing
  
  selectLine = SheetList.SelectedItem.Text
  
  If BK_sheetSheetList.Range("E" & selectLine) = "��" And BK_sheetSheetList.Range("F" & selectLine).Value <> "" Then
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
      .ColumnHeaders.add , "_Display", "�\��", 30
      .ColumnHeaders.add , "_SheetName", "�V�[�g��", 140
    
    For line = 1 To endLine
      With .ListItems.add
        .Text = BK_sheetSheetList.Range("D" & line).Value
        .SubItems(1) = BK_sheetSheetList.Range("E" & line).Value
        .SubItems(2) = BK_sheetSheetList.Range("F" & line).Value
      End With
    Next
    '�ŏI�s�ɋ󔒒ǉ�
      With .ListItems.add
        .Text = line
        .SubItems(1) = ""
        .SubItems(2) = "�V�[�g����"
      End With
    
    .ListItems(selectLine).EnsureVisible
    .ListItems(selectLine).Selected = True
    .SetFocus
  End With
  
End Function

