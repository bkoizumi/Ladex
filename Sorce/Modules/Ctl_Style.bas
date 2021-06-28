Attribute VB_Name = "Ctl_Style"
'**************************************************************************************************
' * �X�^�C��Import/Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Export()
  Dim filePath As String, fileName As String
  Dim FSO As Object
     
     
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.Export"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------

  BK_sheetStyle.Copy
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  With ActiveWorkbook
    With FSO
      fileName = thisAppName & "_" & .GetBaseName(.GetTempName) & ".xlsx"
      filePath = .GetSpecialFolder(2) & "\" & fileName
    End With
    .SaveAs filePath
  End With
  Set FSO = Nothing
  
  Call Ctl_SaveVal.setVal("ExportStyleFilePaht", filePath)
  Call Ctl_SaveVal.setVal("ExportStyleFileName", fileName)


  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'==================================================================================================
Function Import()
  Dim FSO As Object
  Dim filePath As String, fileName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
     
     
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Style.Import"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  targetFilePath = Ctl_SaveVal.getVal("ExportStyleFilePaht")
  Set targetBook = Workbooks.Open(targetFilePath)
  
  targetBook.Sheets("Style").Columns("A:J").Copy BK_ThisBook.Worksheets("Style").Range("A1")
  
  Call Ctl_SaveVal.delVal("ExportStyleFilePaht")
  Call Ctl_SaveVal.delVal("ExportStyleFileName")


  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function

