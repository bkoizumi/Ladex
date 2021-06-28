Attribute VB_Name = "Ctl_Favorite"
'**************************************************************************************************
' * ���C�ɓ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getList()
  Dim tmp, i As Long, buf As String
  
  Call init.setting
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  If Not IsEmpty(tmp) Then
    For i = 0 To UBound(tmp)
      BK_sheetFavorite.Range("A" & i + 2) = tmp(i, 1)
    Next i
  End If
End Function


'==================================================================================================
Function addList()

  Dim line As Long, endLine As Long
  
  Call init.setting
  endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  Call Library.delRegistry("FavoriteList")
  
  For line = 2 To endLine
    Call Library.setRegistry("FavoriteList", "Favorite" & line - 1, BK_sheetFavorite.Range("A" & line))
  Next
End Function


'==================================================================================================
'���C�ɓ���ǉ�
Function add(Optional filePath As String)

  Dim line As Long, endLine As Long
  
  Call init.setting
  line = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row + 1
  
  If filePath = "" Then
    filePath = ActiveWorkbook.FullName
  End If
  BK_sheetFavorite.Range("A" & line) = filePath
  
  ThisWorkbook.Save
'  Set targetBook = Workbooks("�����e�i���X�p.xlsx")
'  ThisWorkbook.Worksheets("Favorite").Columns("A:C").Copy targetBook.Sheets("Favorite").Range("A1")

End Function


'==================================================================================================
'�ڍו\��
Function detail()
  Dim line As Long, endLine As Long
  Dim regLists As Variant
  
  On Error GoTo catchError
  Call getList
  topPosition = Library.getRegistry("UserForm", "FavoriteTop")
  leftPosition = Library.getRegistry("UserForm", "FavoriteLeft")
  
  With Frm_Favorite
    If topPosition = 0 Then
      .StartUpPosition = 2
    Else
      .StartUpPosition = 0
      .Top = topPosition
      .Left = leftPosition
    End If
    .Show vbModeless
  End With
  Call RefreshListBox
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function



'==================================================================================================
Function RefreshListBox()
  Dim line As Long, endLine As Long
  Dim FSO As Object
  
  Call init.setting
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  Frm_Favorite.Lst_Favorite.Clear
  For line = 2 To endLine
    Frm_Favorite.Lst_Favorite.AddItem FSO.GetBaseName(BK_sheetFavorite.Range("A" & line))
  Next
  Set FSO = Nothing
  
  ThisWorkbook.Save
End Function


'**************************************************************************************************
' * �t�H�[����ł̉E�N���b�N
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
  
  BK_sheetFavorite.Range("A" & line).Cut
  BK_sheetFavorite.Range("A" & 2).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveUp()
  Dim line As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  
  BK_sheetFavorite.Range("A" & line).Cut
  BK_sheetFavorite.Range("A" & line - 1).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveDown()
  Dim line As Long, endLine As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  BK_sheetFavorite.Range("A" & line).Cut
  BK_sheetFavorite.Range("A" & line + 1).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function moveBottom()
  Dim line As Long, endLine As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  BK_sheetFavorite.Range("A" & line).Cut
  BK_sheetFavorite.Range("A" & endLine).Insert Shift:=xlDown
  
  Call RefreshListBox
End Function


'==================================================================================================
Function delete()
  Dim line As Long
  Dim filePath As String
  
  Call init.setting
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  
  BK_sheetFavorite.Rows(line & ":" & line).delete Shift:=xlUp
  
  Call RefreshListBox
End Function









