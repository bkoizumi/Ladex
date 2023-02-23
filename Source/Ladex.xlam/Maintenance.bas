Attribute VB_Name = "Maintenance"
Option Explicit

Function レジストリ登録情報抽出()
  Dim tmpRegList
  Dim line
  
  Call init.setting
  Call Library.startScript

  tmpRegList = GetAllSettings(thisAppName, "Main")
  For line = 0 To UBound(tmpRegList)
    Cells(line + 3, 7) = "Main"
    Cells(line + 3, 8) = tmpRegList(line, 0)
    Cells(line + 3, 9) = tmpRegList(line, 1)
  Next
  
  Call Library.endScript(True)
  Call init.unsetting
End Function
