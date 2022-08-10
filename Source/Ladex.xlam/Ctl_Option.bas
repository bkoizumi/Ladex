Attribute VB_Name = "Ctl_Option"
Option Explicit

'**************************************************************************************************
' * オプション画面表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showVersion()
'  On Error GoTo catchError
  
  Call init.setting
  With Frm_Version
    .Label1.Caption = "Ladex Addin For Excel Library"
    .Label2.Caption = "Ver " & thisAppVersion
    .Label3.Caption = "当ソフトはフリーソフトウェアです。" & vbNewLine & _
                      "個人・法人に限らず利用者は自由に使用および配布することができますが、著作権は作者にあります。" & vbNewLine & _
                      "当ソフトを使用した事によるいかなる損害も作者は一切の責任を負いません｡" & vbNewLine & _
                      "ソースを利用する場合にはMITライセンスです｡"
    .Show
  End With
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function showHelp()
  Const funcName As String = "Ctl_Option.showHelp"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  BK_ThisBook.Activate
  LadexSh_Help.Activate
  Sheets("Help").copy
  ActiveWindow.DisplayGridlines = False
  Set targetBook = ActiveWorkbook
  
  With targetBook.VBProject
    .VBComponents.Import (LadexDir & "\RibbonSrc\Ctl_Help.bas")
  End With
  
  'マクロ埋め込み-----------------------------------------------------------------------
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
    .InsertLines 3, "  call Ctl_Help.目次生成"
    .InsertLines 4, ""
    .InsertLines 5, "End Sub"
  End With
  
  targetBook.Activate
  Set targetBook = Nothing
  
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
Function initialization()
  Dim RegistryKey As String, RegistrySubKey As String, RegistryVal As String
  Dim line As Long, endLine As Long
  Dim regName As String

  Const funcName As String = "Ctl_Option.initialization"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  BK_ThisBook.Activate
  endLine = LadexSh_Config.Cells(Rows.count, Library.getColumnNo(BK_setVal("Cells_RegistryKey"))).End(xlUp).Row
  
  Call Library.delRegistry("Main")
  For line = 3 To endLine
    RegistryKey = LadexSh_Config.Range(BK_setVal("Cells_RegistryKey") & line)
    RegistrySubKey = LadexSh_Config.Range(BK_setVal("Cells_RegistrySubKey") & line)
    RegistryVal = LadexSh_Config.Range(BK_setVal("Cells_RegistryValue") & line)
    
    If RegistryKey <> "" Then
     Call Library.setRegistry(RegistryKey, RegistrySubKey, RegistryVal)
    End If
  Next
  Call Ctl_Hollyday.InitializeHollyday

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
Function showOption()
  Dim line As Long, endLine As Long
  Const funcName As String = "Ctl_Option.showOption"

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
  '----------------------------------------------
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 0
    '.Show vbModeless
    .Show
  End With

  ThisWorkbook.Save
  Call init.setting(True)
  Call Main.setShortcutKey


  Exit Function

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
Function HighLight()
  Const funcName As String = "Ctl_Option.HighLight"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 1
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page3.Visible = False
    .MultiPage1.Page4.Visible = False
    .MultiPage1.Page5.Visible = False
    .Show
  End With

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
Function Comment()
  Const funcName As String = "Ctl_Option.Comment"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  With Frm_Option
    .MultiPage1.SelectedItem.Index = 2
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page2.Visible = False
    .MultiPage1.Page4.Visible = False
    .MultiPage1.Page5.Visible = False
    .Show
  End With

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

