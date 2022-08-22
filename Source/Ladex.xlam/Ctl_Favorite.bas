Attribute VB_Name = "Ctl_Favorite"
Option Explicit



'**************************************************************************************************
' * お気に入り
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
Function getList()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, cateEndLine As Long
  Dim tmp, i As Long, buf As String

  Call init.setting
  Call Library.delSheetData(LadexSh_Favorite)
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  
  colLine = 1
  line = 1
  If Not IsEmpty(tmp) Then
    'カテゴリー抽出------------------------------
    For i = 0 To UBound(tmp)
      LadexSh_Favorite.Range("A" & line) = Split(tmp(i, 0), "<L|>")(0)
      line = line + 1
    Next
    
    '重複削除
    endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
    If endLine > 1 Then
      LadexSh_Favorite.Range("A1:A" & endLine).RemoveDuplicates Columns:=1, Header:=xlNo
    End If
    endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row

    colLine = 2
    For line = 1 To endLine
      For i = 0 To UBound(tmp)
        If tmp(i, 0) Like LadexSh_Favorite.Range("A" & line) & "<L|>*" Then
          cateEndLine = LadexSh_Favorite.Cells(Rows.count, colLine).End(xlUp).Row
          If LadexSh_Favorite.Cells(cateEndLine, colLine) <> "" Then
            cateEndLine = cateEndLine + 1
          End If
          LadexSh_Favorite.Cells(cateEndLine, colLine) = tmp(i, 1)
        End If
      Next
      cateEndLine = 1
      colLine = colLine + 1
    Next
  End If
End Function


'==================================================================================================
Function addList()
  Dim line As Long, endLine As Long
  Dim cateLine As Long, cateEndLine As Long
  
  Call init.setting
  cateEndLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row

  Call Library.delRegistry("FavoriteList", "")
  
  If LadexSh_Favorite.Range("A1").Text <> "" And LadexSh_Favorite.Range("B1").Text <> "" Then
    For cateLine = 1 To cateEndLine
      For line = 1 To LadexSh_Favorite.Cells(Rows.count, cateLine + 1).End(xlUp).Row
        Call Library.setRegistry("FavoriteList", LadexSh_Favorite.Range("A" & cateLine) & "<L|>" & line - 1, LadexSh_Favorite.Cells(line, cateLine + 1))
      Next
    Next
  End If
End Function


'==================================================================================================
'お気に入り追加
Function add(Optional setCategory As Long = 1, Optional filePath As String)
  Dim line As Long, endLine As Long
  
  Const funcName As String = "Ctl_Favorite.add"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  If setCategory = -1 Then
    endLine = 2
    setCategory = 0
    LadexSh_Favorite.Range("A1") = "Category01"
  Else
    endLine = LadexSh_Favorite.Cells(Rows.count, setCategory + 1).End(xlUp).Row
  End If
  If endLine <> 1 Then
    endLine = endLine + 1
  End If
  
  If filePath = "" Then
    filePath = ActiveWorkbook.FullName
  End If
  
  Call Library.showDebugForm("filePath", filePath, "debug")
  Call Library.showDebugForm("Cells", Cells(endLine, setCategory + 1).Address, "debug")
  LadexSh_Favorite.Cells(endLine, setCategory + 1) = filePath
  
  Call Library.setRegistry("targetInfo", "FavoriteDirPath", ActiveWorkbook.path)
  Call addList

  '処理終了--------------------------------------
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
'詳細表示
Function detail()
  Dim line As Long, endLine As Long
  Dim regLists As Variant
  Dim topPosition As Long, leftPosition As Long
  
  On Error GoTo catchError
  If Workbooks.count = 0 Then
    Call MsgBox("ブックが開かれていません", vbCritical, thisAppName)
    Exit Function
  End If
  
  Call Ctl_Favorite.getList
  With Frm_Favorite
    .Show vbModeless
  End With
'  Call Ctl_Favorite.RefreshListBox_new
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'**************************************************************************************************
' * フォーム上での右クリック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function moveTop()
  Dim line As Long, colLine As Long
  Dim filePath As String
  
  If Frm_Favorite.Lst_Favorite.ListIndex = 0 Then
    Exit Function
  End If
  
  Call init.setting
  Call Library.startScript
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  LadexSh_Favorite.Cells(line, colLine).Cut
  LadexSh_Favorite.Cells(1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
  
End Function


'==================================================================================================
Function moveUp()
  Dim line As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  
  LadexSh_Favorite.Cells(line, colLine).Cut
  LadexSh_Favorite.Cells(line - 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
End Function


'==================================================================================================
Function moveDown()
  Dim line As Long, endLine As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  endLine = LadexSh_Favorite.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  LadexSh_Favorite.Cells(line, colLine).Cut
  LadexSh_Favorite.Cells(line + 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
  Frm_Favorite.Lst_Favorite.ListIndex = line - 1
  
End Function


'==================================================================================================
Function moveBottom()
  Dim line As Long, endLine As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 1
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  endLine = LadexSh_Favorite.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  LadexSh_Favorite.Cells(line, colLine).Cut
  LadexSh_Favorite.Cells(endLine + 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
  
End Function


'==================================================================================================
Function delete()
  Dim line As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 1
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  LadexSh_Favorite.Cells(line, colLine + 1).delete Shift:=xlUp
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1

End Function


