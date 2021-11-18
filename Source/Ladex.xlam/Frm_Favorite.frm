VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Favorite 
   Caption         =   "お気に入り一覧"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "Frm_Favorite.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Favorite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myMenu As Variant




'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  
  
  Set myMenu = Application.CommandBars.add(Position:=msoBarPopup, Temporary:=True)
  
  With myMenu
    With .Controls.add
      .Caption = "先頭に移動"
      .OnAction = "Ctl_Favorite.moveTop"
      .FaceId = 594
    End With
    With .Controls.add
      .Caption = "1つ上に移動"
      .OnAction = "Ctl_Favorite.moveUp"
      .FaceId = 595
    End With
    With .Controls.add
      .BeginGroup = True
      .Caption = "1つ下に移動"
      .OnAction = "Ctl_Favorite.moveDown"
      .FaceId = 596
    End With
    With .Controls.add
      .Caption = "最後に移動"
      .OnAction = "Ctl_Favorite.moveBottom"
      .FaceId = 597
    End With
    With .Controls.add
      .BeginGroup = True
      .Caption = "削除"
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
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'閉じる
Private Sub Submit_Click()

  Call Library.setRegistry("UserForm", "FavoriteTop", Me.Top)
  Call Library.setRegistry("UserForm", "FavoriteLeft", Me.Left)
  
  Call Ctl_Favorite.addList
  Unload Me
End Sub


'==================================================================================================
'追加
Private Sub add_Click()
  'Dim wsh As Object
  Dim filePath As String
  
  'Set wsh = CreateObject("Wscript.Shell")
  'filePath = Library.getFilePath(wsh.SpecialFolders("Desktop"), "", "お気に入りに追加するファイル", 1)
  filePath = Library.getFilePath("C:", "", "お気に入りに追加するファイル", 1)
  If filePath = "" Then
    End
  End If
  Call Ctl_Favorite.add(filePath)
  Call Ctl_Favorite.RefreshListBox
  
  Set wsh = Nothing
  
End Sub


'==================================================================================================
'リストボックス
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
    
  
    DetailMeg = "<<ファイル情報>>" & vbNewLine
    DetailMeg = DetailMeg & "パ　ス：" & filePath & vbNewLine
    DetailMeg = DetailMeg & "作成日：" & Format(fileInfo.DateCreated, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "更新日：" & Format(fileInfo.DateLastModified, "yyyy/mm/dd hh:nn:ss") & vbNewLine
    DetailMeg = DetailMeg & "サイズ：" & Library.convscale(fileInfo.Size) & " [" & Format(fileInfo.Size, "#,##0") & " Byte" & "]" & vbNewLine
    
    DetailMeg = DetailMeg & "種　類：" & fileInfo.Type
  Else
    DetailMeg = "<<ファイル情報>>" & vbNewLine
    DetailMeg = DetailMeg & "ファイルが存在しません"

  End If
 
  Frm_Favorite.DetailMeg.Value = DetailMeg
  Set FSO = Nothing
  
  Exit Sub
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Sub


