Attribute VB_Name = "Ctl_Window"
Option Explicit

Function 画面サイズ変更(Optional widthVal As Long, Optional heightVal As Long)
  Dim actWin As Window
  Dim actWinTop
  Dim actWinLeft

  Call Library.startScript
  Set actWin = Application.Windows(ActiveWorkbook.Name)
  
  actWinTop = actWin.Top
  actWinLeft = actWin.Left

  Call Library.showDebugForm("actWinTop ：" & actWinTop)
  Call Library.showDebugForm("actWinLeft：" & actWinLeft)

  Call Library.getMachineInfo
  Call Library.showDebugForm("appWidth ：" & MachineInfo("appWidth"))
  Call Library.showDebugForm("appHeight ：" & MachineInfo("appHeight"))
  
  
  actWin.WindowState = xlNormal
  actWin.Height = heightVal
  actWin.Width = widthVal
  
  actWin.Top = actWinTop
  actWin.Left = actWinLeft
  Set actWin = Nothing
  
  Call Library.endScript
End Function
