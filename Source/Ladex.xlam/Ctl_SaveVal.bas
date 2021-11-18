Attribute VB_Name = "Ctl_SaveVal"
Option Explicit

'**************************************************************************************************
' * VBA実行前の値を保持
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setVal(pType As String, pText As String)
  Dim line As Long, endLine As Long
  Dim chkFlg As Boolean
  
  '処理開始--------------------------------------
'  On Error GoTo catchError
  funcName = "Ctl_SaveVal.setVal"

  Call init.setting
  chkFlg = False
  BK_ThisBook.Activate
'  Set BK_sheetSetting = ActiveWorkbook.Worksheets("設定")
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
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function getVal(pType As String) As String
  Dim resetObjVal          As Object
  Dim line As Long, endLine As Long
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  funcName = "Ctl_SaveVal.setVal"

  Call init.setting
'  Set BK_sheetSetting = ActiveWorkbook.Worksheets("設定")
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
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Set resetObjVal = Nothing
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

'==================================================================================================
Function delVal(pType As String)
  Dim line As Long, endLine As Long
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  funcName = "Ctl_SaveVal.setVal"
  
  Call Library.startScript
  Call init.setting
'  Set BK_sheetSetting = ActiveWorkbook.Worksheets("設定")
  '----------------------------------------------
  
  
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) Like "*" & pType Then
      BK_sheetSetting.Range(BK_setVal("Cells_pType") & line) = ""
      BK_sheetSetting.Range(BK_setVal("Cells_pText") & line) = ""

    End If
  Next
 
  
  Call Library.endScript
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

