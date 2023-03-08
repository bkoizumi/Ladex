Attribute VB_Name = "Ctl_File"
Option Explicit

'**************************************************************************************************
' * �t�@�C���A�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �t�@�C���p�X���(Optional dirPath As String = "", Optional line As Long)
  Dim endLine As Long, colLine As Long
  Dim objFolder As Folder
  Dim objFile As File
  Const funcName As String = "Ctl_File.�t�@�C���p�X���"
  
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
  
  If dirPath = "" Then
    With Frm_GetFile
      .Caption = "�t�@�C���p�X���"
      .Show
    End With

    line = 0
    dirPath = FrmVal("targetDir01")
    
    If dirPath = "" Then
      Call Library.errorHandle
    End If
      Call Ctl_ProgressBar.showStart
  End If
  
  With CreateObject("Scripting.FileSystemObject")
    If FrmVal("getSubDir01") = True Then
      For Each objFolder In .GetFolder(dirPath).SubFolders
        If FrmVal("getFullPath01") = True Then
          ActiveCell.Offset(line) = objFolder.path
        Else
          ActiveCell.Offset(line) = objFolder.Name
        End If
         Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, 1, 10, "�t�@�C���p�X���F" & objFolder.Name)
        
        colLine = 1
        '�쐬��
        If FrmVal("getCreateAt01") = True Then
          ActiveCell.Offset(line, colLine) = Format(.GetFolder(objFolder).DateCreated, "yyyy/mm/dd hh:nn:ss")
          ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        End If
        
        '�X�V��
        If FrmVal("getUpdateAt01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = Format(.GetFolder(objFolder).DateLastModified, "yyyy/mm/dd hh:nn:ss")
          ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        End If
        
        '�g���q
        If FrmVal("getExtension01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = "Dir"
        End If
        
        '�T�C�Y
        If FrmVal("getSize01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = Library.convscale(objFolder.Size)
          ActiveCell.Offset(line, colLine).HorizontalAlignment = xlRight
        End If
        line = line + 1
        Call Ctl_File.�t�@�C���p�X���(objFolder.path, line)
      Next
    End If
    For Each objFile In .GetFolder(dirPath).Files
      If FrmVal("getFullPath01") = True Then
        ActiveCell.Offset(line) = objFile.path
      Else
        ActiveCell.Offset(line) = objFile.Name
      End If
      
      
      colLine = 1
      '�쐬��
      If FrmVal("getCreateAt01") = True Then
        ActiveCell.Offset(line, colLine) = Format(.GetFile(objFile).DateCreated, "yyyy/mm/dd hh:nn:ss")
        ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
      End If
      
      '�X�V��
      If FrmVal("getUpdateAt01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = Format(.GetFile(objFile).DateLastModified, "yyyy/mm/dd hh:nn:ss")
        ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
      End If
      
      '�g���q
      If FrmVal("getExtension01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = .GetExtensionName(objFile)
      End If
      
      '�T�C�Y
      If FrmVal("getSize01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = Library.convscale(.GetFile(objFile).Size)
        ActiveCell.Offset(line, colLine).HorizontalAlignment = xlRight
      End If
      
      Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, 1, 10, "�t�@�C���p�X���F" & objFile.Name)
      line = line + 1
    Next
  End With
  

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function �t�H���_����()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Dim basePath As String, targetDir As String
  Dim targetFile As File
  Dim FSO As New FileSystemObject
  
  
  Const funcName As String = "Ctl_File.�t�H���_����"

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
  slctCellsCnt = 0
  basePath = ""
  
  For Each slctCells In Selection
    targetDir = slctCells.Value
    If targetDir <> "" Then
      Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, slctCellsCnt, Selection.CountLarge, "�t�H���_�����F" & targetDir)
      
      If targetDir Like "[A-z]:\*" Then
        Call Library.showDebugForm("targetDir", "�t���p�X", "debug")
        
      ElseIf targetDir Like "\\*" Then
        Call Library.showDebugForm("targetDir", "�l�b�g���[�N�h���C�u", "debug")
      
      Else
        If basePath = "" Then
          basePath = Library.getDirPath(ThisWorkbook.path, "�e�t�H���_�[�̑I��")
        End If
        
        If basePath = "" Then
          Call Library.showNotice(100, , True)
        End If
        If targetDir Like "[\,/]*" Then
          targetDir = basePath & targetDir
        Else
          targetDir = basePath & "\" & targetDir
        End If
      End If
      targetDir = Replace(targetDir, "/", "\")
      
      Call Library.showDebugForm("targetDir", targetDir, "debug")
      Call Library.execMkdir(targetDir)
      
      slctCellsCnt = slctCellsCnt + 1
      DoEvents
      
    End If
  Next

  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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
Function �摜�\�t��()
  Dim line As Long, endLine As Long
  Dim imgFile As Variant
  Dim fileShape As Shape
  Dim fileInfo As Object
  Dim chfkFlg As Boolean
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_File.�摜�\�t��"
  
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
  PrgP_Max = 2
  chfkFlg = False
  '----------------------------------------------
  
  topPosition = ActiveCell.Top
  leftPosition = ActiveCell.Left
  
  For Each imgFile In Library.getFilesPath(ActiveWorkbook.path, "�摜", "img", "pasteImgPath")
    If imgFile <> "" Then
      chfkFlg = True
      Call Library.showDebugForm("imgFile", imgFile, "debug")
      Call Library.getFileInfo(CStr(imgFile), fileInfo)
      
      Set fileShape = ActiveSheet.Shapes.AddPicture( _
        fileName:=imgFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=leftPosition, _
        Top:=topPosition, _
        Width:=0, _
        Height:=0)
      
      fileShape.Name = "Ladex_" & fileInfo("fileName")
      fileShape.LockAspectRatio = msoTrue
      
      '���{�ŕ\��
      fileShape.ScaleWidth 1, msoTrue
      fileShape.ScaleHeight 1, msoTrue
      
      '�g���ݒ�
      With fileShape.line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(127, 127, 127)
        .Transparency = 0
      End With
      
      topPosition = topPosition + fileShape.Height + 20
      Set fileShape = Nothing
    End If
  Next
  
  '�I�u�W�F�N�g�I�����[�h�ɂ���
'  If chfkFlg = True Then
'    With Application.CommandBars.FindControl(ID:=182)
'      If .State = msoButtonUp Then .Execute
'    End With
'  End If
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
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


