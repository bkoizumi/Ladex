Attribute VB_Name = "Ctl_Favorite"
Option Explicit

Const moduleDebug     As Boolean = False

'**************************************************************************************************
' * ���C�ɓ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
Function getList()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, cateEndLine As Long
  Dim tmp, i As Long, buf As String

  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  Call init.setting
  Call Library.delSheetData(targetSheet)
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Call Library.Sort_QuickSort(tmp, LBound(tmp), UBound(tmp), 0)
  colLine = 1
  line = 1
  If Not IsEmpty(tmp) Then
    '�J�e�S���[���o------------------------------
    For i = 0 To UBound(tmp)
      targetSheet.Range("A" & line) = Split(tmp(i, 0), "<L|>")(0)
      line = line + 1
    Next
    
    '�d���폜
    endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
    If endLine > 1 Then
      targetSheet.Range("A1:A" & endLine).RemoveDuplicates Columns:=1, Header:=xlNo
    End If
    endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row

    colLine = 2
    For line = 1 To endLine
      For i = 0 To UBound(tmp)
        If tmp(i, 0) Like targetSheet.Range("A" & line) & "<L|>*" Then
          cateEndLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
          If targetSheet.Cells(cateEndLine, colLine) <> "" Then
            cateEndLine = cateEndLine + 1
          End If
          targetSheet.Cells(cateEndLine, colLine) = tmp(i, 1)
        End If
      Next
      cateEndLine = 1
      colLine = colLine + 1
    Next
    
    If moduleDebug = True Then
      targetSheet.Cells.EntireColumn.AutoFit
      targetSheet.Range("A1").Select
    End If
  End If
End Function


'==================================================================================================
Function ���W�X�g���o�^()
  Dim line As Long, endLine As Long
  Dim cateLine As Long, cateEndLine As Long
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  Call init.setting
  cateEndLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
  Call Library.delRegistry("FavoriteList", "")
  
  If targetSheet.Range("A1").Text <> "" And targetSheet.Range("B1").Text <> "" Then
    For cateLine = 1 To cateEndLine
      For line = 1 To targetSheet.Cells(Rows.count, cateLine + 1).End(xlUp).Row
        Call Library.setRegistry("FavoriteList", targetSheet.Range("A" & cateLine) & "<L|>" & line - 1, targetSheet.Cells(line, cateLine + 1))
      Next
    Next
  End If
End Function


'==================================================================================================
'���C�ɓ���ǉ�
Function add(Optional setCategory As Long = 1, Optional filePath As String)
  Dim line As Long, endLine As Long
  
  Const funcName As String = "Ctl_Favorite.add"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  If setCategory = -1 Then
    endLine = 2
    setCategory = 0
    targetSheet.Range("A1") = "Category01"
  Else
    endLine = targetSheet.Cells(Rows.count, setCategory + 1).End(xlUp).Row
  End If
  If endLine <> 1 Then
    endLine = endLine + 1
  End If
  
  If filePath = "" Then
    filePath = ActiveWorkbook.FullName
  End If
  
  Call Library.showDebugForm("filePath", filePath, "debug")
  Call Library.showDebugForm("Cells", Cells(endLine, setCategory + 1).Address, "debug")
  targetSheet.Cells(endLine, setCategory + 1) = filePath
  
  Call Library.setRegistry("targetInfo", "FavoriteDirPath", ActiveWorkbook.path)
  Call ���W�X�g���o�^

  '�����I��--------------------------------------
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
'�ڍו\��
Function detail()
  Dim line As Long, endLine As Long
  Dim regLists As Variant
  Dim topPosition As Long, leftPosition As Long
  
  On Error GoTo catchError
  If Workbooks.count = 0 Then
    Call MsgBox("�u�b�N���J����Ă��܂���", vbCritical, thisAppName)
    Exit Function
  End If
  
  Call Ctl_Favorite.getList
  With Frm_Favorite
    .Show vbModeless
  End With
'  Call Ctl_Favorite.RefreshListBox_new
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
    Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'**************************************************************************************************
' * �t�H�[����ł̉E�N���b�N
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
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
  
End Function


'==================================================================================================
Function moveUp()
  Dim line As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(line - 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
End Function


'==================================================================================================
Function moveDown()
  Dim line As Long, endLine As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(line + 1, colLine).Insert Shift:=xlDown
  
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
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 1
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 2
  End If
  
  endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(endLine + 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 2
  
End Function


'==================================================================================================
Function delete()
  Dim line As Long, colLine As Long
  Dim filePath As String
  
  Call init.setting
  Call Library.startScript
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 1
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 1
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  targetSheet.Cells(line, colLine + 1).delete Shift:=xlUp
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1

End Function


