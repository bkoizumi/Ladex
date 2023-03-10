Attribute VB_Name = "Ctl_Favorite"
Option Explicit


'**************************************************************************************************
' * ���C�ɓ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
Function chkDebugMode()
  If favoriteDebug = True Then
    Set targetSheet = Workbooks("�����e�i���X�p.xlsm").Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If

End Function


'==================================================================================================
Function ���X�g�擾()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, cateEndLine As Long
  Dim tmp, i As Long, buf As String
  Dim categoryName As String, oldCategoryName As String, FilePath As String

  Const funcName As String = "Ctl_Favorite.���X�g�擾"
  
  '�����J�n--------------------------------------
  On Error Resume Next
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
  Else
    Call init.setting
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Ctl_Favorite.chkDebugMode
  Call Library.delSheetData(targetSheet)
  '----------------------------------------------
  
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  If Not IsEmpty(tmp) Then
    Call Library.Sort_QuickSort(tmp, LBound(tmp), UBound(tmp), 0)
    colLine = 1
    line = 1
    oldCategoryName = ""
    
    For i = 0 To UBound(tmp)
      categoryName = Split(tmp(i, 0), "<L|>")(0)
      FilePath = tmp(i, 1)
      
      Call Library.showDebugForm("categoryName", categoryName, "debug")
      Call Library.showDebugForm("FilePath    ", FilePath, "debug")
      
      If oldCategoryName <> categoryName And oldCategoryName <> "" Then
        colLine = colLine + 1
      End If
      targetSheet.Cells(1, colLine) = categoryName
      oldCategoryName = categoryName
      
      endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row + 1
      targetSheet.Cells(endLine, colLine) = FilePath
    
    Next
    
    If favoriteDebug = True Then
      targetSheet.Cells.EntireColumn.AutoFit
      targetSheet.Range("A1").Select
    End If
  End If
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ���W�X�g���o�^()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim categoryName As String, FilePath As String
  
  Const funcName As String = "Ctl_Favorite.���W�X�g���o�^"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  Call Library.delRegistry("FavoriteList", "")
  
  endColLine = targetSheet.Cells(1, Columns.count).End(xlToLeft).Column
  For colLine = 1 To endColLine
    endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
    For line = 2 To endLine
      categoryName = targetSheet.Cells(1, colLine) & "<L|>" & line - 2
      FilePath = targetSheet.Cells(line, colLine)
    
      Call Library.showDebugForm("FavoriteList", categoryName & "," & FilePath, "debug")
      
      Call Library.setRegistry("FavoriteList", categoryName, FilePath)
    Next
  Next
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function �ǉ�(Optional setCategory As Long = 1, Optional FilePath As String)
  Dim line As Long, endLine As Long
  
  Const funcName As String = "Ctl_Favorite.�ǉ�"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  Call Ctl_Favorite.chkDebugMode
  
  If setCategory = -1 Then
    endLine = 2
    setCategory = 0
    targetSheet.Range("A1") = "Category01"
  Else
    endLine = targetSheet.Cells(Rows.count, setCategory).End(xlUp).Row + 1
  End If
  
  If FilePath = "" Then
    FilePath = ActiveWorkbook.FullName
  End If
  
  Call Library.showDebugForm("setCategory", setCategory, "debug")
  Call Library.showDebugForm("filePath   ", FilePath, "debug")
  Call Library.showDebugForm("Cells      ", Cells(endLine, setCategory).Address, "debug")
  
  Cells(endLine, setCategory).Select
  targetSheet.Cells(endLine, setCategory) = FilePath
  
  Call Library.setRegistry("targetInfo", "FavoriteDirPath", Library.getFileInfo(FilePath, , "CurrentDir"))
  Call Ctl_Favorite.���W�X�g���o�^

  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------
  
  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Function


'==================================================================================================
Function �ڍו\��()
  Dim line As Long, endLine As Long
  Dim regLists As Variant
  Dim topPosition As Long, leftPosition As Long
  
  Const funcName As String = "Ctl_Favorite.�ڍו\��"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  On Error GoTo catchError
  If Workbooks.count = 0 Then
    Call MsgBox("�u�b�N���J����Ă��܂���", vbCritical, thisAppName)
    Exit Function
  End If
  
  Call Ctl_Favorite.���X�g�擾
  With Frm_Favorite
    .Show vbModeless
  End With
  
  '�����I��--------------------------------------
  Exit Function
  '----------------------------------------------
  
  '�G���[������------------------------------------------------------------------------------------
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
  Dim FilePath As String
  
  Const funcName As String = "Ctl_Favorite.moveTop"
  
  '�����J�n--------------------------------------
  If Frm_Favorite.Lst_Favorite.ListIndex = 0 Then
    Exit Function
  End If
  
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(2, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1
  Frm_Favorite.Lst_Favorite.ListIndex = 0
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function moveUp()
  Dim line As Long, colLine As Long
  Dim FilePath As String
  
  Const funcName As String = "Ctl_Favorite.moveUp"
  
  '�����J�n--------------------------------------
  If Frm_Favorite.Lst_Favorite.ListIndex = 0 Then
    Exit Function
  End If
  
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(line - 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1
  Frm_Favorite.Lst_Favorite.ListIndex = line - 3
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function moveDown()
  Dim line As Long, endLine As Long, colLine As Long
  Dim FilePath As String
  
  Const funcName As String = "Ctl_Favorite.moveDown"
  
  '�����J�n--------------------------------------
  If Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1 Then
    Exit Function
  End If
  
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
'  targetSheet.Cells(line, colLine).Select
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(line + 2, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1
  Frm_Favorite.Lst_Favorite.ListIndex = line - 1
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function moveBottom()
  Dim line As Long, endLine As Long, colLine As Long
  Dim FilePath As String
  
  Const funcName As String = "Ctl_Favorite.moveBottom"
  
  '�����J�n--------------------------------------
  If Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1 Then
    Exit Function
  End If
  
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
  endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
  
  If line >= endLine Then
    Exit Function
  End If
  'targetSheet.Cells(line, colLine).Select
  targetSheet.Cells(line, colLine).Cut
  targetSheet.Cells(endLine + 1, colLine).Insert Shift:=xlDown
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1
  Frm_Favorite.Lst_Favorite.ListIndex = endLine - 2
  
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function delete()
  Dim line As Long, colLine As Long
  Dim FilePath As String
  
  Const funcName As String = "Ctl_Favorite.delete"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  line = Frm_Favorite.Lst_Favorite.ListIndex + 2
  If Frm_Favorite.Lst_FavCategory.ListIndex = -1 Then
    colLine = 2
  Else
    colLine = Frm_Favorite.Lst_FavCategory.ListIndex + 1
  End If
  
'  targetSheet.Cells(line, colLine).Select
  targetSheet.Cells(line, colLine).delete Shift:=xlUp
  
  Call Frm_Favorite.RefreshListBox
  Frm_Favorite.Lst_FavCategory.ListIndex = colLine - 1
  Frm_Favorite.Lst_Favorite.ListIndex = 0

  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function



