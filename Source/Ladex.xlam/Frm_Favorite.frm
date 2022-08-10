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
Dim arrFavCategory()

Const addCategoryVal    As String = "≪カテゴリー追加≫"

'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "[" & thisAppName & "] お気に入り"
  
'  Erase arrFavCategory
  
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
  
  Call Frm_Favorite.RefreshListBox
End Sub


'==================================================================================================
Private Sub Lst_Favorite_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
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

'  Call Library.setRegistry("UserForm", "FavoriteTop", Me.Top)
'  Call Library.setRegistry("UserForm", "FavoriteLeft", Me.Left)

  Call Ctl_Favorite.addList
'  Call Library.delSheetData(LadexSh_Favorite)
  Unload Me
  Call init.unsetting
End Sub


'==================================================================================================
'追加
Private Sub add_Click()
  Dim filePath As String
  
  filePath = Library.getFilePath("C:", "", "お気に入りに追加するファイル", 1)
  If filePath <> "" Then
    Call Ctl_Favorite.add(Lst_FavCategory.ListIndex, filePath)
    Call Frm_Favorite.RefreshListBox
  End If
  
End Sub


'==================================================================================================
'リストボックス
Private Sub Lst_Favorite_Click()
  Dim DetailMeg As String
  Dim favLine As Long, catLine As Long
  Dim filePath As String
  
  Dim FSO As Object, fileInfo As Object
  On Error GoTo catchError
  
  Call init.setting
  
  catLine = Lst_FavCategory.ListIndex + 1
  favLine = Lst_Favorite.ListIndex + 1
  
  
  
  filePath = arrFavCategory(catLine, favLine)
  
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
'==================================================================================================
'リストボックス
Private Sub Lst_Favorite_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim filePath As String
  
  If Lst_FavCategory.ListIndex < 0 Then
    Frm_Favorite.DetailMeg.Value = "登録するカテゴリーを選択してください"
  
  ElseIf Lst_FavCategory.list(Lst_FavCategory.ListIndex, 0) = addCategoryVal Then
    Frm_Favorite.DetailMeg.Value = "カテゴリーを登録・選択してください"
  Else
    filePath = Library.getFilePath("C:", "", "お気に入りに追加するファイル", 1)
    If filePath <> "" Then
      Call Ctl_Favorite.add(Lst_FavCategory.ListIndex, filePath)
      Call Frm_Favorite.RefreshListBox
    End If
  End If
  
  Exit Sub
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Sub


'==================================================================================================
'カテゴリー用リストボックス
Private Sub Lst_FavCategory_Click()
  Dim DetailMeg As String
  Dim line As Long, y As Long
  Dim filePath As String
  
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
    Frm_Favorite.DetailMeg.Value = "ダブルクリックでカテゴリーを追加できます"
  Else
    Frm_Favorite.DetailMeg.Value = "ダブルクリックでカテゴリー名変更が可能"
  End If
  

  
  Exit Sub
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Sub

'==================================================================================================
'カテゴリー用リストボックス
Private Sub Lst_FavCategory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim line As Long
  Dim newCategoryName As String
  
  
  line = Lst_FavCategory.ListIndex
  If line = -1 Then
    newCategoryName = InputBox("カテゴリー名を入力してください", "カテゴリー名入力", "")
  Else
    newCategoryName = InputBox("カテゴリー名を入力してください", "カテゴリー名入力", Lst_FavCategory.list(line))
  End If
  
  
  If newCategoryName <> "" Then
    If line = -1 Then
      endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
      LadexSh_Favorite.Range("A" & endLine + 1) = newCategoryName
      
    Else
      LadexSh_Favorite.Range("A" & line + 1) = newCategoryName
    End If
  End If
  Call Frm_Favorite.RefreshListBox
End Sub



'==================================================================================================
Function RefreshListBox()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim line2 As Long, oldEndLine As Long
  
'  Dim LadexSh_Favorite As Worksheet
'  Set LadexSh_Favorite = ThisWorkbook.Worksheets("Favorite")
  
  Const funcName As String = "Frm_Favorite.RefreshListBox"
  
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  
  
  Erase arrFavCategory
  endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
  
  Frm_Favorite.Lst_FavCategory.Clear
  Frm_Favorite.Lst_Favorite.Clear
  
  'カテゴリーリスト生成
  If LadexSh_Favorite.Range("A1") <> "" Then
    For line = 1 To LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
      Call Library.showDebugForm("Lst_FavCategory", LadexSh_Favorite.Range("A" & line), "debug")
      Frm_Favorite.Lst_FavCategory.AddItem LadexSh_Favorite.Range("A" & line)
    Next
  End If
  
  
  '配列の要素数確認------------------------------
  endColLine = LadexSh_Favorite.Cells(1, Columns.count).End(xlToLeft).Column
  oldEndLine = 1
  For colLine = 2 To endColLine
    endLine = LadexSh_Favorite.Cells(Rows.count, colLine).End(xlUp).Row
    If oldEndLine < endLine Then
      oldEndLine = endLine
    End If
  Next
  ReDim Preserve arrFavCategory(1 To endColLine, 0 To oldEndLine)
  
  For colLine = 2 To endColLine
    endLine = LadexSh_Favorite.Cells(Rows.count, colLine).End(xlUp).Row
    arrFavCategory(colLine - 1, 0) = LadexSh_Favorite.Range("A" & colLine - 1)
    
    For line = 1 To endLine
      arrFavCategory(colLine - 1, line) = LadexSh_Favorite.Cells(line, colLine)
      
      If colLine = 2 Then
        Frm_Favorite.Lst_Favorite.AddItem Library.getFileInfo(LadexSh_Favorite.Cells(line, colLine), , "fileName")
      End If
    Next
  Next
  
  Frm_Favorite.Lst_FavCategory.AddItem addCategoryVal
  
  
  
  Call Library.endScript
  'ThisWorkbook.Save
  
End Function

