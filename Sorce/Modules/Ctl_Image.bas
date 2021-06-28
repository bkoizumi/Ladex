Attribute VB_Name = "Ctl_Image"

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
  
  '処理開始--------------------------------------
'  On Error GoTo catchError

  Call init.setting
'  Call Library.startScript
  '----------------------------------------------

  If IsMissing(defSlctArea) Then
    imageName = thisAppName & "ExportImg_" & Format(Now(), "yyyymmdd_hhnnss") & ".png"
    saveDir = Library.getDirPath(ActiveWorkbook.Path, "画像")
    Set slctArea = Selection
  Else
'    imageName = thisAppName & "ExportPreviewImg" & ".jpg"
    saveDir = LadexDir
    Set slctArea = defSlctArea
  End If
  
  If saveDir = "" Then
    Call Library.showNotice(200, "", True)
  End If
  
  If TypeName(slctArea) = "Range" Then
    slctArea.CopyPicture Appearance:=xlScreen, Format:=xlPicture
  
  ElseIf TypeName(slctArea) = "ChartArea" Then
    slctArea.Copy
  End If
  
  ActiveWorkbook.Activate
  ActiveSheet.Select
'  Call Library.waitTime(1000)
  
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
  Call Library.endScript
  If IsMissing(defSlctArea) Then
    Call Shell("Explorer.exe /select, " & saveDir & "\" & imageName, vbNormalFocus)
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, "", True)
End Function

