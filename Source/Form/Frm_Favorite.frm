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




'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
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
End Sub


'==================================================================================================
Private Sub Lst_Favorite_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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

  Call Library.setRegistry("UserForm", "FavoriteTop", Me.Top)
  Call Library.setRegistry("UserForm", "FavoriteLeft", Me.Left)
  
  Call Ctl_Favorite.addList
  Unload Me
End Sub


'==================================================================================================
'�ǉ�
Private Sub add_Click()
  'Dim wsh As Object
  Dim filePath As String
  
  'Set wsh = CreateObject("Wscript.Shell")
  'filePath = Library.getFilePath(wsh.SpecialFolders("Desktop"), "", "���C�ɓ���ɒǉ�����t�@�C��", 1)
  filePath = Library.getFilePath("C:", "", "���C�ɓ���ɒǉ�����t�@�C��", 1)
  If filePath = "" Then
    End
  End If
  Call Ctl_Favorite.add(filePath)
  Call Ctl_Favorite.RefreshListBox
  
  Set wsh = Nothing
  
End Sub


'==================================================================================================
'���X�g�{�b�N�X
Private Sub Lst_Favorite_Click()
  Dim DetailMeg As String
  Dim line As Long
  Dim filePath As String
  
  Dim FSO As Object, fileInfo As Object
  On Error GoTo catchError
  
  Call init.setting
  
  line = Lst_Favorite.ListIndex + 2
  filePath = BK_sheetFavorite.Range("A" & line)
  
  If Library.chkFileExists(filePath) Then
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fileInfo = FSO.GetFile(filePath)
    
  
    DetailMeg = "<<�t�@�C�����>>" & vbNewLine
    DetailMeg = DetailMeg & "�p�@�X�F" & filePath & vbNewLine
    DetailMeg = DetailMeg & "�쐬���F" & Format(fileInfo.DateCreated, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "�X�V���F" & Format(fileInfo.DateLastModified, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "�T�C�Y�F" & Format(fileInfo.Size, "#,##0") & " Byte" & vbNewLine
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


