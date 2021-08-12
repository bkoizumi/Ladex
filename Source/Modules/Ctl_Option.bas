Attribute VB_Name = "Ctl_Option"
Option Explicit

'**************************************************************************************************
' * �I�v�V������ʕ\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showHelp()
  
  Call init.setting(True)
  BK_ThisBook.Activate
  BK_sheetHelp.Activate
  Sheets("Help").Copy
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
  
End Function


'==================================================================================================
Function initialization()
  Dim RegistryKey As String, RegistrySubKey As String, RegistryVal As String
  Dim line As Long, endLine As Long
  Dim regName As String

  Call init.setting(True)
  
  BK_ThisBook.Activate
  endLine = BK_sheetSetting.Cells(Rows.count, Library.getColumnNo(BK_setVal("Cells_RegistryKey"))).End(xlUp).Row
  
  Call Library.delRegistry("Main")
  Call Library.delRegistry("UserForm")

  
  For line = 3 To endLine
    RegistryKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistryKey") & line)
    RegistrySubKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistrySubKey") & line)
    RegistryVal = BK_sheetSetting.Range(BK_setVal("Cells_RegistryValue") & line)
    
    If RegistryKey <> "" Then
     Call Library.setRegistry(RegistryKey, RegistrySubKey, RegistryVal)
    End If
  Next
  
  Call Ctl_Hollyday.InitializeHollyday


End Function

'==================================================================================================
Function showOption()
  Dim topPosition As Long, leftPosition As Long
'  On Error GoTo catchError
  
  Call init.setting(True)
  topPosition = Library.getRegistry("UserForm", "OptionTop")
  leftPosition = Library.getRegistry("UserForm", "OptionLeft")
  
  Call Ctl_UsrForm.�\���ʒu(topPosition, leftPosition)
  With Frm_Option
    .StartUpPosition = 0
    .Top = topPosition
    .Left = leftPosition
    .MultiPage1.SelectedItem.Index = 0
'    .MultiPage1.Page1.Visible = True
'    .MultiPage1.Page2.Visible = True
'    .MultiPage1.Page3.Visible = True
    
    '.Show vbModeless
    .Show
  End With

  Exit Function

'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function HighLight()
  Dim topPosition As Long, leftPosition As Long
  On Error GoTo catchError
  
  topPosition = Library.getRegistry("UserForm", "OptionTop")
  leftPosition = Library.getRegistry("UserForm", "OptionLeft")
  
  Call Ctl_UsrForm.�\���ʒu(topPosition, leftPosition)
  With Frm_Option
    .StartUpPosition = 0
    .Top = topPosition
    .Left = leftPosition
    .MultiPage1.SelectedItem.Index = 1
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page3.Visible = False
    
    .Show
  End With

  Exit Function

'�G���[������=====================================================================================
catchError:
  'Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'==================================================================================================
Function Comment()
  Dim topPosition As Long, leftPosition As Long
  On Error GoTo catchError
  
  topPosition = Library.getRegistry("UserForm", "OptionTop")
  leftPosition = Library.getRegistry("UserForm", "OptionLeft")
  
  Call Ctl_UsrForm.�\���ʒu(topPosition, leftPosition)
  With Frm_Option
    .StartUpPosition = 0
    .Top = topPosition
    .Left = leftPosition
    .MultiPage1.SelectedItem.Index = 2
    .MultiPage1.Page1.Visible = False
    .MultiPage1.Page2.Visible = False
    
    .Show
  End With

  Exit Function

'�G���[������=====================================================================================
catchError:
  'Call Library.showNotice(Err.Number, Err.Description, True)
End Function

