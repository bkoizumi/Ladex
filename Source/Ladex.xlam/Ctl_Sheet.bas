Attribute VB_Name = "Ctl_Sheet"
Option Explicit

'**************************************************************************************************
' * R1C1�\�L
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function R1C1�\�L()

  On Error Resume Next
  
  Call init.setting
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
  Else
    Application.ReferenceStyle = xlA1
  End If
  
End Function

'**************************************************************************************************
' * �Z�����E��������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �Z��������()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  
  Const funcName As String = "Ctl_Sheet.�Z��������"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Cells.EntireColumn.AutoFit
  
  If IsNumeric(Range("A1").Text) Then
    Call Library.setColumnWidth
  Else
    For colLine = 1 To Columns.count
      If Cells(1, colLine).ColumnWidth > 30 Then
        colName = Library.getColumnName(colLine)
        Columns(colName & ":" & colName).ColumnWidth = 30
      End If
    Next
  End If
  Call Library.endScript(True)
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function

'==================================================================================================
Function �Z����������()
  Call Library.startScript
  Call init.setting
  
  Cells.EntireRow.AutoFit
  Call Library.endScript(True)
End Function











