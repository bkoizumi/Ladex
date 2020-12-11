VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Favorite 
   Caption         =   "お気に入り"
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
' * 処理キャンセル
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()
  Unload Me
End Sub




'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  Unload Me
End Sub


'**************************************************************************************************
' * リストボックス
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub Lst_Favorite_Click()
  Dim DetailMeg As String
  Dim line As Long
  Dim filePath As String
  
  On Error GoTo catchError
  
  Call init.setting
  line = Lst_Favorite.ListIndex + 2
  filePath = sheetFavorite.Range("A" & line)
  
  DetailMeg = "<<ファイル情報>>" & vbNewLine
  DetailMeg = DetailMeg & "パス　：" & filePath & vbNewLine
  DetailMeg = DetailMeg & "更新日：" & FileDateTime(filePath) & vbNewLine
 
 
  Frm_Favorite.DetailMeg.Value = DetailMeg
  
  Exit Sub
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Sub


