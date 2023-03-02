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

Const addCategoryVal  As String = "≪カテゴリー追加≫"
Const moduleDebug     As Boolean = False

'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()

  Call Ctl_Favorite.chkDebugMode
  
  
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
  
  Frm_Favorite.Lst_FavCategory.ListIndex = 0
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
  Call Ctl_Favorite.レジストリ登録
  
  Call Ctl_Favorite.chkDebugMode
  Call Library.delSheetData(targetSheet)
  Unload Me
End Sub


'==================================================================================================
'追加
Private Sub add_Click()
  Dim FilePath As String
  Dim categoryIndex As Variant
  
  categoryIndex = Frm_Favorite.Lst_FavCategory.ListIndex
  
  FilePath = Library.getRegistry("targetInfo", "FavoriteDirPath")
  If FilePath = "" Then
    FilePath = "C:"
  End If
  
  FilePath = Library.getFilePath(FilePath, "", "お気に入りに追加するファイル", 1)
  If FilePath <> "" Then
    Call Ctl_Favorite.追加(Lst_FavCategory.ListIndex + 1, FilePath)
    Call Frm_Favorite.RefreshListBox
    
    Frm_Favorite.Lst_FavCategory.ListIndex = categoryIndex
    Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1
  
  End If
  
End Sub


'==================================================================================================
'リストボックス
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
    
  
    DetailMeg = "<<ファイル情報>>" & vbNewLine
    DetailMeg = DetailMeg & "パ　ス：" & FilePath & vbNewLine
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
  Dim FilePath As String
  Dim categoryIndex As Variant
  
  If Lst_FavCategory.ListIndex < 0 Then
    Frm_Favorite.DetailMeg.Value = "登録するカテゴリーを選択してください"
    MsgBox "登録するカテゴリーを選択してください", vbExclamation
  
  ElseIf Lst_FavCategory.list(Lst_FavCategory.ListIndex, 0) = addCategoryVal Then
    Frm_Favorite.DetailMeg.Value = "カテゴリーの登録・選択してください"
    MsgBox "カテゴリーの登録・選択してください", vbExclamation
    
    
  Else
    FilePath = Library.getFilePath("C:", "", "お気に入りに追加するファイル", 1)
    If FilePath <> "" Then
      categoryIndex = Frm_Favorite.Lst_FavCategory.ListIndex
      Call Ctl_Favorite.追加(Lst_FavCategory.ListIndex + 1, FilePath)
      Call Frm_Favorite.RefreshListBox
      
    Frm_Favorite.Lst_FavCategory.ListIndex = categoryIndex
    Frm_Favorite.Lst_Favorite.ListIndex = Frm_Favorite.Lst_Favorite.ListCount - 1
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
  
  If moduleDebug = True Then
    Set targetSheet = ActiveWorkbook.Worksheets("Favorite")
  Else
    Set targetSheet = ThisWorkbook.Worksheets("Favorite")
  End If
    
    
  
  If newCategoryName <> "" Then
    '重複チェック
    endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
    If WorksheetFunction.CountIf(targetSheet.Range("A1:A" & endLine), newCategoryName) > 1 Then
      Frm_Favorite.DetailMeg.Value = "登録するカテゴリーが重複しています"
      MsgBox "登録するカテゴリーが重複しています", vbExclamation
    
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
  
  '処理開始--------------------------------------
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
  'カテゴリーリスト生成--------------------------
  For colLine = 1 To endColLine
    categoryName = targetSheet.Cells(1, colLine)
    Call Library.showDebugForm("categoryName", categoryName, "debug")
    
    Frm_Favorite.Lst_FavCategory.AddItem categoryName
  Next
  
  
  '配列の要素数確認------------------------------
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

  '最後に≪カテゴリー追加≫を追加
  Frm_Favorite.Lst_FavCategory.AddItem addCategoryVal
  
  '処理終了--------------------------------------
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end1")
  Exit Function
  '----------------------------------------------

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
