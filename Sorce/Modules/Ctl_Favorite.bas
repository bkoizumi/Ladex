Attribute VB_Name = "Ctl_Favorite"
'**************************************************************************************************
' * お気に入り
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************



'==================================================================================================
'お気に入り追加
Function add(Optional filePath As String)

  Dim line As Long, endLine As Long
  
  Call init.setting
  line = sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row + 1
  
  If filePath = "" Then
    filePath = ActiveWorkbook.FullName
  End If
  sheetFavorite.Range("A" & line) = filePath
  
'  Set targetBook = Workbooks("メンテナンス用.xlsx")
'  ThisWorkbook.Worksheets("Favorite").Columns("A:C").Copy targetBook.Sheets("Favorite").Range("A1")

End Function


'==================================================================================================
'詳細表示
Function detail()
  Dim line As Long, endLine As Long
  Dim regLists As Variant
  
'  On Error GoTo catchError
  
  With Frm_Favorite
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 4)
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Show vbModeless
  End With
  Call RefreshListBox
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function



'==================================================================================================
Function RefreshListBox()
  Dim line As Long, endLine As Long
  Dim FSO As Object
  
  Call init.setting
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  endLine = sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  Frm_Favorite.Lst_Favorite.Clear
  For line = 2 To endLine
    Frm_Favorite.Lst_Favorite.AddItem FSO.GetBaseName(sheetFavorite.Range("A" & line))
  Next
  Set FSO = Nothing
End Function


'**************************************************************************************************
' * フォーム上での右クリック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function moveTop()
  Dim line As Long
  Dim filePath As String
  
  If Frm_Favorite.Lst_Favorite.ListIndex = 0 Then
    Exit Function
  End If
  
  Call init.setting
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  
  sheetFavorite.Range("A" & line).Cut
  sheetFavorite.Range("A" & 2).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveUp()
  Dim line As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  
  sheetFavorite.Range("A" & line).Cut
  sheetFavorite.Range("A" & line - 1).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveDown()
  Dim line As Long, endLine As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  endLine = sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  sheetFavorite.Range("A" & line).Cut
  sheetFavorite.Range("A" & line + 1).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveBottom()
  Dim line As Long, endLine As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  endLine = sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  sheetFavorite.Range("A" & line).Cut
  sheetFavorite.Range("A" & endLine).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function delete()
  Dim line As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  
  sheetFavorite.Rows(line & ":" & line).delete Shift:=xlUp
  
  Call RefreshListBox
End Function









