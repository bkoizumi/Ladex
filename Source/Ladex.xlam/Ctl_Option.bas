Attribute VB_Name = "Ctl_Option"
Option Explicit

'**************************************************************************************************
' * �I�v�V������ʕ\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showVersion()
  Dim strBuf As String
'  On Error GoTo catchError
  
  Call init.setting
  
  strBuf = strBuf & "Ladex Addin For Excel Library Ver. " & thisAppVersion & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " Copyright (c) 2021 Bunpei.Koizumi" & vbCrLf
  strBuf = strBuf & " author:bunpei.koizumi@gmail.com" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " The MIT License (MIT)" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " Permission is hereby granted, free of charge, to any person obtaining a copy" & vbCrLf
  strBuf = strBuf & " of this software and associated documentation files (the ""Software""), to deal" & vbCrLf
  strBuf = strBuf & " in the Software without restriction, including without limitation the rights" & vbCrLf
  strBuf = strBuf & " to use, copy, modify, merge, publish, distribute, sublicense, and/or sell" & vbCrLf
  strBuf = strBuf & " copies of the Software, and to permit persons to whom the Software is" & vbCrLf
  strBuf = strBuf & " furnished to do so, subject to the following conditions:" & vbCrLf
  strBuf = strBuf & "" & vbCrLf
  strBuf = strBuf & " The above copyright notice and this permission notice shall be included in all" & vbCrLf
  strBuf = strBuf & " copies or substantial portions of the Software." & vbCrLf
  
  
  With Frm_Version
    .Label1.Caption = "Ladex Addin For Excel Library"
    .Label2.Caption = "Ver " & thisAppVersion
    .TextBox1.Value = "���\�t�g�̓t���[�\�t�g�E�F�A�ł��B" & vbNewLine & _
                      "�l�E�@�l�Ɍ��炸���p�҂͎��R�Ɏg�p����єz�z���邱�Ƃ��ł��܂����A���쌠�͍�҂ɂ���܂��B" & vbNewLine & _
                      "���\�t�g���g�p�������ɂ�邢���Ȃ鑹�Q����҂͈�؂̐ӔC�𕉂��܂���" & vbNewLine & _
                      "�\�[�X�𗘗p����ꍇ�ɂ�MIT���C�Z���X�ł��" & vbNewLine & vbNewLine & strBuf
                      
                      
    .Show
  End With
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function showHelp()
  Const funcName As String = "Ctl_Option.showHelp"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  LadexBook.Activate
  LadexSh_Help.Activate
  Sheets("Help").copy
  ActiveWindow.DisplayGridlines = False
  Set targetBook = ActiveWorkbook
  
  With targetBook.VBProject
    .VBComponents.Import (LadexDir & "\RibbonSrc\Ctl_Help.bas")
  End With
  
  '�}�N�����ߍ���-----------------------------------------------------------------------
  With targetBook.VBProject.VBComponents.Item("Help").CodeModule
    .InsertLines 1, "Private Sub Worksheet_SelectionChange(ByVal Target As Range)"
    .InsertLines 2, ""
    .InsertLines 3, "  On Error GoTo catchError"
    .InsertLines 4, "  If ActiveCell.Column = 1 And ActiveCell.Value <> """" Then"
    .InsertLines 5, "    With ActiveWindow"
    .InsertLines 6, "      .ScrollRow = Target.Row"
    .InsertLines 7, "      .ScrollColumn = Target.Column"
    .InsertLines 8, "    End With"
    .InsertLines 9, "  End If"
    .InsertLines 10, "Exit Sub"
    .InsertLines 11, "catchError:"
    .InsertLines 12, ""
    .InsertLines 13, ""
    .InsertLines 14, "End Sub"
  End With
  
  With targetBook.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
    .InsertLines 1, "Private Sub Workbook_Activate()"
    .InsertLines 2, ""
    .InsertLines 3, "  call Ctl_Help.�ڎ�����"
    .InsertLines 4, ""
    .InsertLines 5, "End Sub"
  End With
  
  targetBook.Activate
  Set targetBook = Nothing
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function initialization()
  Dim RegistryKey As String, RegistrySubKey As String, RegistryVal As String
  Dim line As Long, endLine As Long
  Dim regName As String

  Const funcName As String = "Ctl_Option.initialization"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  LadexBook.Activate
  endLine = LadexSh_Config.Cells(Rows.count, Library.getColumnNo(dicVal("Cells_RegistryKey"))).End(xlUp).Row
  
  Call Library.delRegistry("Main")
  For line = 3 To endLine
    RegistryKey = LadexSh_Config.Range(dicVal("Cells_RegistryKey") & line)
    RegistrySubKey = LadexSh_Config.Range(dicVal("Cells_RegistrySubKey") & line)
    RegistryVal = LadexSh_Config.Range(dicVal("Cells_RegistryValue") & line)
    
    If RegistryKey <> "" Then
     Call Library.setRegistry(RegistryKey, RegistrySubKey, RegistryVal)
    End If
  Next
  Call Ctl_Hollyday.InitializeHollyday

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function showOption()
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_Option.showOption"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 0
    .Show
  End With

  ThisWorkbook.Save
  Call init.setting(True)
  Call Main.�V���[�g�J�b�g�L�[�ݒ�


  Exit Function

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function HighLight()
  Const funcName As String = "Ctl_Option.HighLight"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 1
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page3.Visible = False
    .MultiPage1.Page4.Visible = False
    .MultiPage1.Page5.Visible = False
    .Show
  End With

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



'==================================================================================================
Function Comment()
  Const funcName As String = "Ctl_Option.Comment"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 2
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page2.Visible = False
    .MultiPage1.Page4.Visible = False
    .MultiPage1.Page5.Visible = False
    .Show
  End With

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function Addin����()
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_Option.Addin����"

  '�����J�n--------------------------------------
'  If runFlg = False Then
'    Call init.setting
'    Call Library.showDebugForm(funcName, , "start1")
'    Call Library.startScript
'  Else
'    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
'  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  LadexSh_Function.Activate
  Workbooks(ThisWorkbook.Name).IsAddin = False


  Exit Function

  '�����I��--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
'    Call init.unsetting
'  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ������()
  Dim line As Long, endLine As Long
  
  Const funcName As String = "Ctl_Option.������"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  Call Library.delRegistry("FavoriteList", "")
  'Call Library.setRegistry("FavoriteList", "Category01<L|>0", "")
  
  'Call Library.delRegistry("Main", "")
  
  endLine = LadexSh_Config.Cells(Rows.count, 7).End(xlUp).Row
  For line = 3 To endLine
    If LadexSh_Config.Range("G" & line) <> "" Then
      Call Library.setRegistry("Main", LadexSh_Config.Range("H" & line).Text, LadexSh_Config.Range("I" & line).Text)
    End If
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �o�b�N�A�b�v()
  Dim line As Long, endLine As Long
  Dim tmpRegList As Variant
  Dim backupFile As String, outMeg As String
  
  Const funcName As String = "Ctl_Option.�o�b�N�A�b�v"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  backupFile = LadexDir & "\Ladex_backup.ini"
  If Library.chkFileExists(backupFile) = True Then
    Call Library.execDel(backupFile)
  End If
  
  outMeg = ""

  Call Library.outputText("[FavoriteList]", backupFile, "utf-8")


  '���C�ɓ��胊�X�g------------------------------
  tmpRegList = GetAllSettings(thisAppName, "FavoriteList")
  For line = 0 To UBound(tmpRegList)
    outMeg = tmpRegList(line, 0) & "=" & tmpRegList(line, 1)
    Call Library.outputText(outMeg, backupFile, "utf-8")
  Next
  Call Library.outputText("", backupFile, "utf-8")

  'Main�ݒ�--------------------------------------
  Call Library.outputText("[Main]", backupFile, "utf-8")
  
  tmpRegList = GetAllSettings(thisAppName, "Main")
  For line = 0 To UBound(tmpRegList)
    outMeg = tmpRegList(line, 0) & "=" & tmpRegList(line, 1)
    Call Library.outputText(outMeg, backupFile, "utf-8")
  Next
  Call Library.outputText("", backupFile, "utf-8")

  'targetInfo�ݒ�--------------------------------------
  Call Library.outputText("[targetInfo]", backupFile, "utf-8")
  
  tmpRegList = GetAllSettings(thisAppName, "targetInfo")
  For line = 0 To UBound(tmpRegList)
    outMeg = tmpRegList(line, 0) & "=" & tmpRegList(line, 1)
    Call Library.outputText(outMeg, backupFile, "utf-8")
  Next
  Call Library.outputText("", backupFile, "utf-8")
  

  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���X�g�A()
  Dim line As Long, endLine As Long
  
  Const funcName As String = "Ctl_Option.���X�g�A"
  
  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  'Call Library.delRegistry("FavoriteList", "")


  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
