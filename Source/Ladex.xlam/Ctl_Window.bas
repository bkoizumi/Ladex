Attribute VB_Name = "Ctl_Window"
Option Explicit

Function ��ʃT�C�Y�ύX(Optional widthVal As Long, Optional heightVal As Long)
  Dim actWin As Window
  Dim actWinTop
  Dim actWinLeft

  Call Library.startScript
  Set actWin = Application.Windows(ActiveWorkbook.Name)
  
  actWinTop = actWin.Top
  actWinLeft = actWin.Left

  Call Library.showDebugForm("actWinTop �F" & actWinTop)
  Call Library.showDebugForm("actWinLeft�F" & actWinLeft)

  Call Library.getMachineInfo
  Call Library.showDebugForm("appWidth �F" & MachineInfo("appWidth"))
  Call Library.showDebugForm("appHeight �F" & MachineInfo("appHeight"))
  
  
  actWin.WindowState = xlNormal
  actWin.Height = heightVal
  actWin.Width = widthVal
  
  actWin.Top = actWinTop
  actWin.Left = actWinLeft
  Set actWin = Nothing
  
  Call Library.endScript
End Function
