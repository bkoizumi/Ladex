Attribute VB_Name = "Ctl_SaveVal"
'**************************************************************************************************
' * VBA���s�O�̒l��ێ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setVal(pType As String, pText As String)
  Dim line As Long, endLine As Long
  Dim chkFlg As Boolean
  
  '�����J�n--------------------------------------
'  On Error GoTo catchError
  FuncName = "Ctl_SaveVal.setVal"

  Call init.setting
  chkFlg = False
  BK_ThisBook.Activate
  Set BK_sheetSetting = ActiveWorkbook.Worksheets("�ݒ�")
  '----------------------------------------------
  endLine = Cells(Rows.count, 4).End(xlUp).Row
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) = pType Then
      BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) = pType
      BK_sheetSetting.Range(BK_setVal("Cells_pText") & line) = pText
      
      chkFlg = True
      Exit For
    End If
  Next
   
  If chkFlg = False Then
    BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) = pType
    BK_sheetSetting.Range(BK_setVal("Cells_pText") & line) = pText
  End If
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'==================================================================================================
Function getVal(pType As String) As String
  Dim resetObjVal          As Object
  
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  FuncName = "Ctl_SaveVal.setVal"

  Call init.setting
  Set BK_sheetSetting = ActiveWorkbook.Worksheets("�ݒ�")
  '----------------------------------------------
  
  Set resetObjVal = Nothing
  Set resetObjVal = CreateObject("Scripting.Dictionary")
  
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) <> "" Then
      resetObjVal.add BK_sheetSetting.Range(BK_setVal("Cells_pType") & line).Text, BK_sheetSetting.Range(BK_setVal("Cells_pText") & line).Text
    End If
  Next
  
  If resetObjVal("reSet" & pType) = "" Then
    getVal = resetObjVal(pType)
  Else
    getVal = resetObjVal("reSet" & pType)
  End If
  Set resetObjVal = Nothing
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Set resetObjVal = Nothing
  Call Library.showNotice(400, FuncName, True)
End Function

'==================================================================================================
Function delVal(pType As String)
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  FuncName = "Ctl_SaveVal.setVal"
  
  Call Library.startScript
  Call init.setting
  Set BK_sheetSetting = ActiveWorkbook.Worksheets("�ݒ�")
  '----------------------------------------------
  
  
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) Like "*" & pType Then
      BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) = ""
      BK_sheetSetting.Range(BK_setVal("Cells_pText") & line) = ""

    End If
  Next
 
  
  Call Library.endScript
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function

