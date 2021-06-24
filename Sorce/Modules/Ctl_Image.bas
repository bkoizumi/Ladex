Attribute VB_Name = "Ctl_Image"

'**************************************************************************************************
' * 画像処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function saveSelectArea2Image()
  Dim slctArea
  Dim targetImg As Chart
  Dim imageName As String, saveDir As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError

  Call init.setting
  Call Library.startScript
  '----------------------------------------------

  imageName = thisAppName & "ExportImg_" & Format(Now(), "yyyymmdd_hhnnss") & ".png"
  saveDir = Library.getDirPath(ActiveWorkbook.Path, "画像")

  If saveDir = "" Then
    Call Library.showNotice(200, "", True)
  End If
  
  Set slctArea = Selection
  
  If TypeName(Selection) = "Range" Then
    slctArea.CopyPicture Appearance:=xlScreen, Format:=xlPicture
  
  ElseIf TypeName(Selection) = "ChartArea" Then
    Selection.Copy
  End If

  Set targetImg = ActiveSheet.ChartObjects.add(0, 0, slctArea.Width, slctArea.Height).Chart
  With targetImg
    .Parent.Select
    .Paste
    .Export saveDir & "\" & imageName
    .Parent.delete
  End With
  
  Set targetImg = Nothing
  Set slctArea = Nothing

  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call Shell("Explorer.exe /select, " & saveDir & "\" & imageName, vbNormalFocus)
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, "", True)
End Function

