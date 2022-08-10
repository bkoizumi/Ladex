Attribute VB_Name = "Ctl_Image"
Option Explicit

'**************************************************************************************************
' * 画像処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function saveSelectArea2Image(Optional defSlctArea As Variant, Optional imageName As Variant)
  Dim slctArea
  Dim targetImg As Chart
  Dim saveDir As String
  Dim DisplayGridlineFlg As Boolean
  Const funcName As String = "Ctl_Image.saveSelectArea2Image"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  DisplayGridlineFlg = False
  '----------------------------------------------

  If IsMissing(defSlctArea) Then
    imageName = thisAppName & "ExportImg_" & Format(Now(), "yyyymmdd_hhnnss") & ".png"
    imageName = "Img_" & Format(Now(), "yyyymmdd_hhnnss") & ".png"
    saveDir = LadexDir & "\Images\"
    Set slctArea = Selection
  Else
'    imageName = thisAppName & "ExportPreviewImg" & ".jpg"
    saveDir = LadexDir & "\RibbonImg\"
    Set slctArea = defSlctArea
  End If
  
  If Library.chkDirExists(saveDir) = "" Then
    Call Library.execMkdir(saveDir)
  End If
  Select Case TypeName(slctArea)
    Case "Range"
  
      '枠線設定切り替え--------------------------
      If ActiveWindow.DisplayGridlines = True Then
        ActiveWindow.DisplayGridlines = False
        DisplayGridlineFlg = True
      End If
      slctArea.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    Case "ChartArea", "Picture", "GroupObject", "TextBox", "Rectangle"
      slctArea.copy
    
    Case Else
      Call Library.showNotice(5, TypeName(slctArea))
      
  End Select
  
  ActiveWorkbook.Activate
  ActiveSheet.Select
'  Call Library.waitTime(1000)
  
  Set targetImg = ActiveSheet.ChartObjects.add(0, 0, slctArea.Width, slctArea.Height).Chart
  With targetImg
    .Parent.Select
    .Paste
    .Export saveDir & imageName
    .Parent.delete
  End With
  
  Set targetImg = Nothing
  Set slctArea = Nothing

  '枠線設定
  If DisplayGridlineFlg = True Then
    ActiveWindow.DisplayGridlines = True
  End If
  
  
  If IsMissing(defSlctArea) Then
    Call Shell("Explorer.exe /select, " & saveDir & imageName, vbNormalFocus)
  End If
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  End If
  '----------------------------------------------
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 貼付け()
  Dim line As Long, endLine As Long
  Dim objFolder As Folder
  Dim objFile As File
  Dim dirPath  As String
  Const funcName As String = "Ctl_Image.画像貼付け"
  
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
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  
  ActiveSheet.Pictures.Insert objFile.path
  
  
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
