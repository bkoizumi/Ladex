Attribute VB_Name = "Ctl_Sheet"

'**************************************************************************************************
' * R1C1表記
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function R1C1表記()

  On Error Resume Next
  
  If Application.ReferenceStyle = xlA1 Then
    Application.ReferenceStyle = xlR1C1
  Else
    Application.ReferenceStyle = xlA1
  End If
  
End Function

'**************************************************************************************************
' * セル幅・高さ調整
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function セル幅調整()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  
  Call Library.startScript
  Cells.EntireColumn.AutoFit
  
  For colLine = 1 To Columns.count
    If Cells(1, colLine).ColumnWidth > 30 Then
      colName = Library.getColumnName(colLine)
      Columns(colName & ":" & colName).ColumnWidth = 30
    End If
  Next
  
  Call Library.endScript(True)
End Function


'==================================================================================================
Function セル高さ調整()
  Call Library.startScript
  Cells.EntireRow.AutoFit
  Call Library.endScript(True)
End Function

