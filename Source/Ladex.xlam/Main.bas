Attribute VB_Name = "Main"
Option Explicit

'���[�N�u�b�N�p�ϐ�------------------------------
'���[�N�V�[�g�p�ϐ�------------------------------
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
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  BK_ThisBook.Activate
  endLine = BK_sheetSetting.Cells(Rows.count, 7).End(xlUp).Row

  If Library.getRegistry("Main", "debugMode", "String") = "" Then
    For line = 3 To endLine
      RegistryKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistryKey") & line)
      RegistrySubKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistrySubKey") & line)
      val = BK_sheetSetting.Range(BK_setVal("Cells_RegistryValue") & line)

      If RegistryKey <> "" Then
       Call Library.setRegistry(RegistryKey, RegistrySubKey, val)
      End If
    Next
  End If
  
  '�Ǝ��֐��ݒ�----------------------------------
  Call Ctl_Hollyday.InitializeHollyday
  Call Ctl_UsrFunction.InitializeUsrFunction
  
  '�V���[�g�J�b�g�L�[�ݒ�------------------------
  Call Main.setShortcutKey


  '�����I��--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm("", , "end1")
  'Call init.unsetting
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
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
Function setShortcutKey()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim keyVal As Variant
  Dim ShortcutKey As String, ShortcutFunc As String
  Const funcName As String = "Main.setShortcutKey"
  
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
  
  endLine = BK_sheetFunction.Cells(Rows.count, 1).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetFunction.Range("E" & line) <> "" Then
      ShortcutKey = ""
      For Each keyVal In Split(BK_sheetFunction.Range("E" & line), "+")
        If keyVal = "Ctrl" Then
          ShortcutKey = "^"
        ElseIf keyVal = "Alt" Then
          ShortcutKey = ShortcutKey & "%"
        ElseIf keyVal = "Shift" Then
          ShortcutKey = ShortcutKey & "+"
        Else
          ShortcutKey = ShortcutKey & keyVal
        End If
      Next
      ShortcutFunc = "Menu.ladex_" & BK_sheetFunction.Range("C" & line)
      Call Library.showDebugForm("ShortcutKey", ShortcutKey, "debug")
      Call Library.showDebugForm("Function   ", ShortcutFunc, "debug")
      
      Call Application.OnKey(ShortcutKey, ShortcutFunc)
    End If
  Next
  
  'Call Application.OnKey("{F1}", "Ctl_Option.showVersion")
  Call Application.OnKey("{F1}", "")


  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end1")
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
' * �n�C���C�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �n�C���C�g()
'  Dim highLightFlg As String
'  Dim highLightArea As String
'
'  Call Library.startScript
'  highLightFlg = Library.getRegistry(ActiveWorkbook.Name, "HighLightFlg")
'
'  If highLightFlg = "" Then
'    Call Library.setLineColor(Selection.Address, True, Library.getRegistry("HighLightColor"))
'
'    Call Library.setRegistry(ActiveWorkbook.Name, True, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightSheet", ActiveSheet.Name, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightArea", Selection.Address, "HighLightFlg")
'
'  Else
'    highLightArea = Library.getRegistry(ActiveWorkbook.Name & "_HighLightArea")
'
'    If highLightArea = "" Then
'      highLightArea = Selection.Address
'    End If
'    Call Library.unsetLineColor(highLightArea)
'
'    Call Library.delRegistry(ActiveWorkbook.Name, "HighLightFlg")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightSheet")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightArea")
'  End If
'
'  Call Library.endScript(True)

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
  
  BK_sheetStyle.Copy
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
  targetBook.Sheets("Style").Columns("A:J").Copy ThisWorkbook.Worksheets("Style").Range("A1")
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

  
  With Application.CommandBars("Cell").Controls.add(before:=1, Type:=msoControlPopup, Temporary:=True)
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









