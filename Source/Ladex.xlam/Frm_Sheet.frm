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
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '�\���ʒu�w��----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True
  
  With Frm_Sheet
    .Caption = "�V�[�g�Ǘ� " & thisAppName
    .inputSheetName.Value = ActiveSheet.Name
    With SheetList
      .View = lvwReport
      .LabelEdit = lvwManual
      .HideSelection = False
      .AllowColumnReorder = True
      .FullRowSelect = True
      .Gridlines = True
      .ColumnHeaders.add , "_ID", "#", 30
      .ColumnHeaders.add , "_Display", "�\��", 30, lvwColumnCenter
      .ColumnHeaders.add , "_SheetName", "�V�[�g��", 140
      
      For line = 1 To ActiveWorkbook.Worksheets.count
        With .ListItems.add
          .Text = line
          
          If ActiveWorkbook.Worksheets(line).Visible = True Then
            .SubItems(1) = "��"
          
          Else
            .SubItems(1) = "�|"
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
  
  InitializeFlg = False
  
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
  Dim SheetName As String, meg As String
  Const funcName As String = "Frm_Sheet.edit_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  SheetName = SheetList.SelectedItem.SubItems(2)
  Frm_Sheet.inputSheetName.Value = SheetName
  
  If ActiveWorkbook.Worksheets(SheetName).Visible = 2 Then
    meg = "�}�N���ɂ���Ĕ�\���ƂȂ��Ă���V�[�g�ł�" & vbNewLine & "�}�N���̓���ɉe����^����\��������܂��B"
    Frm_Sheet.add.Enabled = False
    Frm_Sheet.edit.Enabled = False
    Frm_Sheet.del.Enabled = False
    
    Frm_Sheet.up.Enabled = False
    Frm_Sheet.down.Enabled = False
    
  ElseIf ActiveWorkbook.Worksheets(SheetName).Visible = True Then
    meg = "�_�u���N���b�N�őI��(�A�N�e�B�u��)���܂�"
    Frm_Sheet.add.Enabled = True
    Frm_Sheet.edit.Enabled = True
    Frm_Sheet.del.Enabled = True
  
    Frm_Sheet.up.Enabled = True
    Frm_Sheet.down.Enabled = True
  
  Else
    meg = "��\���ƂȂ��Ă���V�[�g�ł�" & vbNewLine & "�_�u���N���b�N�ŕ\�����A�I��(�A�N�e�B�u��)���܂�"
    Frm_Sheet.add.Enabled = False
    Frm_Sheet.edit.Enabled = False
    Frm_Sheet.del.Enabled = True
    
    Frm_Sheet.up.Enabled = False
    Frm_Sheet.down.Enabled = False
    
  End If

    sheetInfo.Caption = meg & vbNewLine & ""
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
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
    ActiveWorkbook.Sheets(SheetName).Visible = True
    ActiveWorkbook.Worksheets(SheetName).Select
  End If
  
  Unload Me
  Call Library.endScript
End Sub

'==================================================================================================
'��
Private Sub up_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.up_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  Sheets(SheetName).Move Before:=Sheets(SheetList.SelectedItem.Text - 1)
  
  
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub


'==================================================================================================
'��
Private Sub down_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.down_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  Sheets(SheetName).Move Before:=Sheets(SheetList.SelectedItem.Text + 2)
  
  
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'�V�[�g���ύX
Private Sub edit_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.edit_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  If Library.chkSheetExists(inputSheetName.Value) = False Then
    ActiveWorkbook.Sheets(SheetName).Select
    ActiveWorkbook.Sheets(SheetName).Name = inputSheetName.Value
  Else
    sheetInfo.Caption = inputSheetName.Value & "�́A���łɑ��݂��܂�"
  End If
  
  
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'�V�[�g�ǉ�
Private Sub add_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.add_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  ActiveWorkbook.Sheets(SheetName).Select
  
  If Library.chkSheetExists(inputSheetName.Value) = False Then
    Sheets.add After:=ActiveSheet
    ActiveSheet.Name = inputSheetName.Value
  Else
    sheetInfo.Caption = inputSheetName.Value & "�͂��łɑ��݂��܂�"
  End If
  
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub


'==================================================================================================
'�V�[�g�폜
Private Sub del_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.del_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  
'  If MsgBox(SheetName & "���폜���܂�(���ɂ��ǂ��܂���)", vbYesNo + vbExclamation) = vbYes Then
    ActiveWorkbook.Sheets(SheetName).delete
'  End If
  
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub


'==================================================================================================
'�\��/��\��
Private Sub display_Click()
  Dim SheetName As String
   
  Const funcName As String = "Frm_Sheet.display_Click"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  If ActiveWorkbook.Sheets(SheetName).Visible = True Then
    ActiveWorkbook.Sheets(SheetName).Visible = False
  Else
    ActiveWorkbook.Sheets(SheetName).Visible = True
  End If
  Call reLoadList
  
  Exit Sub

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

'==================================================================================================
'�V�[�g�̑I��
Private Sub active_Click()
  Dim SheetName As String, sheetDspFLg As String
  
  SheetName = SheetList.SelectedItem.SubItems(2)
  
  If ActiveWorkbook.Worksheets(SheetName).Visible = True Then
    ActiveWorkbook.Worksheets(SheetName).Select
  Else
    ActiveWorkbook.Worksheets(SheetName).Visible = True
    ActiveWorkbook.Worksheets(SheetName).Select
  End If
  Call reLoadList

End Sub

'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
End Sub

'==================================================================================================
Function reLoadList()
  Dim line As Long, endLine As Long
  Const funcName As String = "Frm_Sheet.reLoadList"

  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
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
    .ColumnHeaders.add , "_Display", "�\��", 30, lvwColumnCenter
    .ColumnHeaders.add , "_SheetName", "�V�[�g��", 140
    
    For line = 1 To ActiveWorkbook.Worksheets.count
      With .ListItems.add
        .Text = line
        
        If ActiveWorkbook.Worksheets(line).Visible = True Then
          .SubItems(1) = "��"
        
        Else
          .SubItems(1) = "�|"
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
End Function

