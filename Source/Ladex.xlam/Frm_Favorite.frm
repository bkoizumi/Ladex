VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Favorite 
   Caption         =   "���C�ɓ���ꗗ"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "Frm_Favorite.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_Favorite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myMenu As Variant
Dim arrFavCategory()

Const addCategoryVal  As String = "��J�e�S���[�ǉ���"
Const moduleDebug     As Boolean = False

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()

  Call Ctl_Favorite.chkDebugMode
  
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "[" & thisAppName & "] ���C�ɓ���"
  
'  Erase arrFavCategory
  
  Set myMenu = Application.CommandBars.add(Position:=msoBarPopup, Temporary:=True)
  
  With myMenu
    With .Controls.add
      .Caption = "�擪�Ɉړ�"
      .OnAction = "Ctl_Favorite.moveTop"
      .FaceId = 594
    End With
    With .Controls.add
      .Caption = "1��Ɉړ�"
      .OnAction = "Ctl_Favorite.moveUp"
      .FaceId = 595
    End With
    With .Controls.add
      .BeginGroup = True
      .Caption = "1���Ɉړ�"
      .OnAction = "Ctl_Favorite.moveDown"
      .FaceId = 596
    End With
    With .Controls.add
      .Caption = "�Ō�Ɉړ�"
      .OnAction = "Ctl_Favorite.moveBottom"
      .FaceId = 597
    End With
    With .Controls.add
      .BeginGroup = True
      .Caption = "�폜"
      .OnAction = "Ctl_Favorite.delete"
      .FaceId = 293
    End With
    
  End With
  
  Call Frm_Favorite.RefreshListBox
  
  Frm_Favorite.Lst_FavCategory.ListIndex = 0
End Sub


'==================================================================================================
Private Sub Lst_Favorite_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button = 2 Then myMenu.ShowPopup
End Sub




'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'����
Private Sub Submit_Click()
  Call Ctl_Favorite.���W�X�g���o�^
  
  Call Ctl_Favorite.chkDebugMode
  Call Library.delSheetData(targetSheet)
  Unload Me
End Sub


'==================================================================================================
'�ǉ�
Private Sub add_Click()
  Dim FilePath As String
  Dim categoryIndex As Variant
  
  categoryIndex = Frm_Favorite.Lst_FavCategory.ListIndex
  
  FilePath = Library.getRegistry("targetInfo", "FavoriteDirPath")
  If FilePath = "" Then
    FilePath = "C:"
  End If
  
  FilePath = Library.getFilePath(FilePath, "", "���C�ɓ���ɒǉ�����t�@�C��", 1)
  If FilePath <> "" Then
    Call Ctl_Favorite.�ǉ�(Lst_FavCategory.ListIndex + 1, FilePath)
    Call Frm_Favorite.RefreshListBox
    
    Frm_Favorite.Lst_FavCategory.ListIndex = categoryIndex
    Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1
  
  End If
  
End Sub


'==================================================================================================
'���X�g�{�b�N�X
Private Sub Lst_Favorite_Click()
  Dim DetailMeg As String
  Dim favLine As Long, catLine As Long
  Dim FilePath As String
  
  Dim FSO As Object, fileInfo As Object
  On Error GoTo catchError
  
  Call init.setting
  
  catLine = Lst_FavCategory.ListIndex + 1
  favLine = Lst_Favorite.ListIndex + 1
  
  
  
  FilePath = arrFavCategory(catLine, favLine)
  
  If Library.chkFileExists(FilePath) Then
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fileInfo = FSO.GetFile(FilePath)
    
  
    DetailMeg = "<<�t�@�C�����>>" & vbNewLine
    DetailMeg = DetailMeg & "�p�@�X�F" & FilePath & vbNewLine
    DetailMeg = DetailMeg & "�쐬���F" & Format(fileInfo.DateCreated, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "�X�V���F" & Format(fileInfo.DateLastModified, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "�T�C�Y�F" & Library.convscale(fileInfo.Size) & " [" & Format(fileInfo.Size, "#,##0") & " Byte" & "]" & vbNewLine
    
    DetailMeg = DetailMeg & "��@�ށF" & fileInfo.Type
  Else
    DetailMeg = "<<�t�@�C�����>>" & vbNewLine
    DetailMeg = DetailMeg & "�t�@�C�������݂��܂���"

  End If
 
  Frm_Favorite.DetailMeg.Value = DetailMeg
  Set FSO = Nothing
  
  Exit Sub
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Sub
'==================================================================================================
'���X�g�{�b�N�X
Private Sub Lst_Favorite_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim FilePath As String
  Dim categoryIndex As Variant
  
  If Lst_FavCategory.ListIndex < 0 Then
    Frm_Favorite.DetailMeg.Value = "�o�^����J�e�S���[��I�����Ă�������"
    MsgBox "�o�^����J�e�S���[��I�����Ă�������", vbExclamation
  
  ElseIf Lst_FavCategory.list(Lst_FavCategory.ListIndex, 0) = addCategoryVal Then
    Frm_Favorite.DetailMeg.Value = "�J�e�S���[�̓o�^�E�I�����Ă�������"
    MsgBox "�J�e�S���[�̓o�^�E�I�����Ă�������", vbExclamation
    
    
  Else
    FilePath = Library.getFilePath("C:", "", "���C�ɓ���ɒǉ�����t�@�C��", 1)
    If FilePath <> "" Then
      categoryIndex = Frm_Favorite.Lst_FavCategory.ListIndex
      Call Ctl_Favorite.�ǉ�(Lst_FavCategory.ListIndex + 1, FilePath)
      Call Frm_Favorite.RefreshListBox
      
    Frm_Favorite.Lst_FavCategory.ListIndex = categoryIndex
    Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1
    End If
  End If
  
  Exit Sub
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Sub


'==================================================================================================
'�J�e�S���[�p���X�g�{�b�N�X
Private Sub Lst_FavCategory_Click()
  Dim DetailMeg As String
  Dim line As Long, y As Long
  Dim FilePath As String
  
  Dim FSO As Object, fileInfo As Object
  On Error GoTo catchError
  
  Call init.setting
  
  line = Lst_FavCategory.ListIndex + 1
  Frm_Favorite.Lst_Favorite.Clear
  
  For y = LBound(arrFavCategory, 2) + 1 To UBound(arrFavCategory, 2)
    If arrFavCategory(line, y) <> "" Then
      Frm_Favorite.Lst_Favorite.AddItem Library.getFileInfo(CStr(arrFavCategory(line, y)), , "fileName")
    End If
  Next
  
  If Lst_FavCategory.list(Lst_FavCategory.ListIndex) = addCategoryVal Then
    Frm_Favorite.DetailMeg.Value = "�_�u���N���b�N�ŃJ�e�S���[��ǉ��ł��܂�"
  Else
    Frm_Favorite.DetailMeg.Value = "�_�u���N���b�N�ŃJ�e�S���[���ύX���\"
  End If
  

  
  Exit Sub
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Sub

'==================================================================================================
'�J�e�S���[�p���X�g�{�b�N�X
Private Sub Lst_FavCategory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim line As Long
  Dim newCategoryName As String
  
  
  line = Lst_FavCategory.ListIndex
  If line = -1 Then
    newCategoryName = InputBox("�J�e�S���[������͂��Ă�������", "�J�e�S���[������", "")
  Else
    newCategoryName = InputBox("�J�e�S���[������͂��Ă�������", "�J�e�S���[������", Lst_FavCategory.list(line))
  End If
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
    
    
  
  If newCategoryName <> "" Then
    '�d���`�F�b�N
    endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
    If WorksheetFunction.CountIf(targetSheet.Range("A1:A" & endLine), newCategoryName) > 1 Then
      Frm_Favorite.DetailMeg.Value = "�o�^����J�e�S���[���d�����Ă��܂�"
      MsgBox "�o�^����J�e�S���[���d�����Ă��܂�", vbExclamation
    
    Else
      If line <> -1 Then
        endLine = line + 1
      End If
      
      targetSheet.Range("A" & endLine) = newCategoryName
    End If
    
  End If
  Call Frm_Favorite.RefreshListBox
End Sub



'==================================================================================================
Function RefreshListBox()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim line2 As Long, oldEndLine As Long
  Dim categoryName As String, FilePath As String
  
  Const funcName As String = "Ctl_Favorite.RefreshListBox"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  Call Library.startScript
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  Erase arrFavCategory
  endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
  
  If endLine = 1 And targetSheet.Range("A1") = "" Then
    targetSheet.Range("A1") = "Category01"
  End If
  
  Frm_Favorite.Lst_FavCategory.Clear
  Frm_Favorite.Lst_Favorite.Clear
  
  endColLine = targetSheet.Cells(1, Columns.count).End(xlToLeft).Column
  '�J�e�S���[���X�g����--------------------------
  For colLine = 1 To endColLine
    categoryName = targetSheet.Cells(1, colLine)
    Call Library.showDebugForm("categoryName", categoryName, "debug")
    
    Frm_Favorite.Lst_FavCategory.AddItem categoryName
  Next
  
  
  '�z��̗v�f���m�F------------------------------
  endColLine = targetSheet.Cells(1, Columns.count).End(xlToLeft).Column
  oldEndLine = 1

  For colLine = 1 To endColLine
    endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
    If oldEndLine < endLine Then
      oldEndLine = endLine
    End If
  Next
  ReDim Preserve arrFavCategory(1 To endColLine, 0 To oldEndLine)

  For colLine = 1 To endColLine
    arrFavCategory(colLine, 0) = targetSheet.Cells(1, colLine).Value
    
    endLine = targetSheet.Cells(Rows.count, colLine).End(xlUp).Row
    For line = 2 To endLine
      arrFavCategory(colLine, line - 1) = targetSheet.Cells(line, colLine).Value
      Frm_Favorite.Lst_Favorite.AddItem Library.getFileInfo(targetSheet.Cells(line, colLine), , "fileName")
    Next
  Next

  '�Ō�Ɂ�J�e�S���[�ǉ����ǉ�
  Frm_Favorite.Lst_FavCategory.AddItem addCategoryVal
  
  '�����I��--------------------------------------
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
