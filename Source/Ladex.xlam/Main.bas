Attribute VB_Name = "Main"
Option Explicit

'���[�N�u�b�N�p�ϐ�------------------------------
''���[�N�V�[�g�p�ϐ�------------------------------
'�O���[�o���ϐ�----------------------------------





'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeBook()
  Dim RegistryKey As String, RegistrySubKey As String, val As String
  Dim line As Long, endLine As Long
  Dim regName As String
  Const funcName As String = "Main.InitializeBook"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  runFlg = True
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  ThisWorkbook.Activate
  endLine = LadexSh_Config.Cells(Rows.count, 7).End(xlUp).Row

  For line = 3 To endLine
    RegistryKey = LadexSh_Config.Range("G" & line)
    RegistrySubKey = LadexSh_Config.Range("H" & line)
    val = LadexSh_Config.Range("I" & line)
    
    If Library.getRegistry(RegistryKey, RegistrySubKey, "String") = "" Then
      Call Library.setRegistry(RegistryKey, RegistrySubKey, val)
    End If
  Next
  
  '�Ǝ��֐��ݒ�----------------------------------
  Call Ctl_Hollyday.InitializeHollyday
  Call Ctl_UsrFunction.InitializeUsrFunction
  
  '�V���[�g�J�b�g�L�[�ݒ�------------------------
  Call Main.�V���[�g�J�b�g�L�[�ݒ�


  '�����I��--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
'**************************************************************************************************
' * �V���[�g�J�b�g�L�[�̐ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �V���[�g�J�b�g�L�[�ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim keyVal As Variant
  Dim ShortcutKey As String, ShortcutFunc As String, ShortcutKey1 As String
  
  Const funcName As String = "Main.�V���[�g�J�b�g�L�[�ݒ�"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call Application.OnKey("{F1}", "")
  
  endLine = LadexSh_Function.Cells(Rows.count, 3).End(xlUp).Row
  For line = 2 To endLine
    If LadexSh_Function.Range("I" & line) <> "" Then
      ShortcutKey = ""
      ShortcutKey1 = ""
      
      For Each keyVal In Split(LadexSh_Function.Range("I" & line), "+")
        If keyVal = "Ctrl" Then
          ShortcutKey = "^"
        ElseIf keyVal = "Alt" Then
          ShortcutKey = ShortcutKey & "%"
        ElseIf keyVal = "Shift" Then
          ShortcutKey = ShortcutKey & "+"
        Else
          Select Case keyVal
            Case 0 To 9
              ShortcutKey1 = ShortcutKey & "{" & 96 + keyVal & "}"
              ShortcutKey = ShortcutKey & keyVal
              
            Case Else
              ShortcutKey = ShortcutKey & keyVal
          End Select
        End If
      Next
      ShortcutFunc = "'Menu.�e�@�\�Ăяo�� """ & LadexSh_Function.Range("E" & line) & """'"
      Call Library.showDebugForm("ShortcutKey ", ShortcutKey, "debug")
      Call Library.showDebugForm("ShortcutFunc", ShortcutFunc, "debug")
      
      Call Application.OnKey(ShortcutKey, ShortcutFunc)
           
      
      
      If ShortcutKey1 <> "" Then
        Call Library.showDebugForm("ShortcutKey", ShortcutKey1, "debug")
        Call Library.showDebugForm("Function   ", ShortcutFunc, "debug")
        Call Application.OnKey(ShortcutKey1, ShortcutFunc)
      End If
    End If
  Next
  
'  Call Application.OnKey("{F1}", "Ctl_Option.showVersion")


  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function xxxxxxxxxx()
End Function

'**************************************************************************************************
' * �摜�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �摜�ݒ�()

  With ActiveWorkbook.ActiveSheet
    Dim AllShapes As Shapes
    Dim CurShape As Shape
    Set AllShapes = .Shapes
    
    For Each CurShape In AllShapes
      CurShape.Placement = xlMove
    Next
  End With
End Function




'**************************************************************************************************
' * �ݒ�Import / Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �ݒ�_���o()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting
  
  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  LadexSh_Style.copy
  ActiveWorkbook.SaveAs fileName:=TempName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  
  Call Library.endScript
  
  MsgBox ("�C��������A�ۑ������Ă�������")
End Function

'==================================================================================================
Function �ݒ�_�捞()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting

  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  Set targetBook = Workbooks.Open(TempName)
  targetBook.Sheets("Style").Columns("A:J").copy ThisWorkbook.Worksheets("Style").Range("A1")
  targetBook.Close
  
  Call FSO.DeleteFile(TempName, True)
  
  Call Ctl_Style.�X�^�C���폜
  Call Library.endScript
End Function

'==================================================================================================
Function �E�N���b�N���j���[(Target As Range, Cancel As Boolean)
  Dim menu01 As CommandBarControl
  
  Call init.setting
  
  '�W����ԂɃ��Z�b�g
  Application.CommandBars("Cell").Reset
  For Each menu01 In Application.CommandBars("Cell").Controls
    'Call Library.showDebugForm("�E�N���b�N", menu01.Caption, "debug")
    
    If menu01.Caption Like "*[�����\�Ƃ��� �ǉ�����]*" Then
      menu01.Visible = False
    End If
  Next

  
  With Application.CommandBars("Cell").Controls.add(Before:=1, Type:=msoControlPopup, Temporary:=True)
    .Caption = thisAppName
    If Not (Target.count = Rows.count Or Target.count = Columns.count) Then
      With .Controls.add(Temporary:=True)
        .Caption = "�s������ւ��ē\�t��"
        .OnAction = "menu.ladex_�s������ւ��ē\�t��"
      End With
    End If
    With .Controls.add(Temporary:=True)
      .BeginGroup = True
      .Caption = "�s�̑}��"
      .OnAction = "menu.ladex_�s�}��"
    End With
    With .Controls.add(Temporary:=True)
      .Caption = "��̑}��"
      .OnAction = "menu.ladex_��}��"
    End With
  End With
  


  Application.CommandBars("Cell").ShowPopup
  Application.CommandBars("Cell").Reset
  Cancel = True
End Function









