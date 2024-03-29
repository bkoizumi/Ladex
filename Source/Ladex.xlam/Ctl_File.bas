Attribute VB_Name = "Ctl_File"
Option Explicit

'==================================================================================================
Function ファイルパス情報(Optional dirPath As String = "", Optional line As Long)
  Dim endLine As Long, colLine As Long
  Dim objFolder As Folder
  Dim objFile As File
  Const funcName As String = "Ctl_File.ファイルパス情報"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  If dirPath = "" Then
    With Frm_GetFile
      .Caption = "ファイルパス情報"
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
         Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, 1, 10, "ファイルパス情報：" & objFolder.Name)
        
        colLine = 1
        '作成日
        If FrmVal("getCreateAt01") = True Then
          ActiveCell.Offset(line, colLine) = Format(.GetFolder(objFolder).DateCreated, "yyyy/mm/dd hh:nn:ss")
          ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        End If
        
        '更新日
        If FrmVal("getUpdateAt01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = Format(.GetFolder(objFolder).DateLastModified, "yyyy/mm/dd hh:nn:ss")
          ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        End If
        
        '拡張子
        If FrmVal("getExtension01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = "Dir"
        End If
        
        'サイズ
        If FrmVal("getSize01") = True Then
          colLine = colLine + 1
          ActiveCell.Offset(line, colLine) = Library.convscale(objFolder.Size)
          ActiveCell.Offset(line, colLine).HorizontalAlignment = xlRight
        End If
        line = line + 1
        Call Ctl_File.ファイルパス情報(objFolder.path, line)
      Next
    End If
    For Each objFile In .GetFolder(dirPath).Files
      If FrmVal("getFullPath01") = True Then
        ActiveCell.Offset(line) = objFile.path
      Else
        ActiveCell.Offset(line) = objFile.Name
      End If
      
      
      colLine = 1
      '作成日
      If FrmVal("getCreateAt01") = True Then
        ActiveCell.Offset(line, colLine) = Format(.GetFile(objFile).DateCreated, "yyyy/mm/dd hh:nn:ss")
        ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
      End If
      
      '更新日
      If FrmVal("getUpdateAt01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = Format(.GetFile(objFile).DateLastModified, "yyyy/mm/dd hh:nn:ss")
        ActiveCell.Offset(line, colLine).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
      End If
      
      '拡張子
      If FrmVal("getExtension01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = .GetExtensionName(objFile)
      End If
      
      'サイズ
      If FrmVal("getSize01") = True Then
        colLine = colLine + 1
        ActiveCell.Offset(line, colLine) = Library.convscale(.GetFile(objFile).Size)
        ActiveCell.Offset(line, colLine).HorizontalAlignment = xlRight
      End If
      
      Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, 1, 10, "ファイルパス情報：" & objFile.Name)
      line = line + 1
    Next
  End With
  

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 画像貼付け()
  Dim line As Long, endLine As Long
  Dim imgFile As Variant
  Dim fileShape As Shape
  Dim fileInfo As Object
  Dim chfkFlg As Boolean
  
  Const funcName As String = "Ctl_File.画像貼付け"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  chfkFlg = False
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------

  For Each imgFile In Library.getFilesPath(ActiveWorkbook.path, "画像", "img", "pasteImgPath")
    If imgFile <> "" Then
      chfkFlg = True
      Call Library.showDebugForm("imgFile", imgFile, "debug")
      Call Library.getFileInfo(CStr(imgFile), fileInfo)
      
      Set fileShape = ActiveSheet.Shapes.AddPicture( _
        fileName:=imgFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=ActiveCell.Left, _
        Top:=ActiveCell.Top, _
        Width:=0, _
        Height:=0)
      
      fileShape.Name = "Ladex_" & fileInfo("fileName")
      fileShape.LockAspectRatio = msoTrue
      
      '等倍で表示
      fileShape.ScaleWidth 1, msoTrue
      fileShape.ScaleHeight 1, msoTrue
      
      '枠線設定
      With fileShape.line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(127, 127, 127)
        .Transparency = 0
      End With
      
      Set fileShape = Nothing
    End If
  Next
  
  'オブジェクト選択モードにする
'  If chfkFlg = True Then
'    With Application.CommandBars.FindControl(ID:=182)
'      If .State = msoButtonUp Then .Execute
'    End With
'  End If
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function フォルダ生成()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim slctCellsCnt As Long
  Dim basePath As String, targetDir As String
  Dim targetFile As File
  Dim FSO As New FileSystemObject
  
  
  Const funcName As String = "Ctl_File.フォルダ生成"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  slctCellsCnt = 0
  basePath = ""
  
  For Each slctCells In Selection
    targetDir = slctCells.Value
    If targetDir <> "" Then
      Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, slctCellsCnt, Selection.CountLarge, "フォルダ生成：" & targetDir)
      
      If targetDir Like "[A-z]:\*" Then
        Call Library.showDebugForm("targetDir", "フルパス", "debug")
        
      ElseIf targetDir Like "\\*" Then
        Call Library.showDebugForm("targetDir", "ネットワークドライブ", "debug")
      
      Else
        If basePath = "" Then
          basePath = Library.getDirPath(ThisWorkbook.path, "親フォルダーの選択")
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

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function
  
'エラー発生時====================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
