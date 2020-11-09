Attribute VB_Name = "GetImage"
'Public path As String
'
'Function GetImage_Paste()
'
'  Dim dirPath As String
'  Dim FolderName As String
'
'  dirPath = Library_GetDirPath("")
'
'  FolderName = Mid(dirPath, InStrRev(dirPath, "\") + 1)
'
'    'ワークシートを追加し、名前付ける
'  Worksheets.Add After:=Worksheets(Worksheets.count)
'  ActiveSheet.Name = FolderName
'
'  ' シートの色を赤色にする
'  With ActiveWorkbook.Sheets(FolderName).Tab
'      .Color = RGB(255, 0, 0)
'      .TintAndShade = 0
'  End With
'
'  Call GetImage_InsertPicture(dirPath, FolderName)
'
'
'
'
'
'
'
'
'
''    Dim dNumber As Long
''    Dim arrDirectory() As String
''
''    'フォルダ選択ダイアログを表示するプロシージャ呼び出す
''    If path = "" Then
''      path = OpenFolderDialog()
''      If path = "" Then
''        Exit Function
''      End If
''    End If
''
''
''    '選択したフォルダ内のフォルダ数を取得するプロシージャを呼び出す
''    dNumber = DirectoryNumber(path)
''
''    ReDim arrDirectory(dNumber)
''
''    'シートを追加するプロシージャを呼び出す
''    arrDirectory = GetImage_AddSheets(path, arrDirectory)
''
''    '画像を挿入するプロシージャを呼び出す
''    Call InsertPicture(dNumber, path, arrDirectory)
''
''    '最初のシート選択
'''    Sheets(arrDirectory(0)).Select
'
'End Function
'
'
''ダイアログを表示するプロシージャ
'Function OpenFolderDialog() As String
'
'    Dim FolderDlg As Office.FileDialog
'
'    'フォルダ選択ダイアログに
'    Set FolderDlg = Application.FileDialog(msoFileDialogFolderPicker)
'    '設定、表示、戻り値
'    With FolderDlg
'      .InitialFileName = ActiveWorkbook.path & "\"
'      .AllowMultiSelect = False
'      If .Show = -1 Then OpenFolderDialog = FolderDlg.SelectedItems(1)
'    End With
'    Set FolderDlg = Nothing
'
'End Function
'
'
''フォルダ数を取得するプロシージャ
'Function DirectoryNumber(argPath As String) As Long
'
'    Dim dName As String
'    Dim num As Long
'
'    '選択したフォルダ名取得
'    dName = Dir(argPath & "\*.*", vbDirectory)
'
'    'フォルダ数ループ
'    Do While dName <> ""
'        'dNameがフォルダかチェック
'        If GetAttr(argPath & "\" & dName) And vbDirectory Then
'            'dNameが、現在フォルダまたは親フォルダかチェック
'            If dName <> "." And dName <> ".." Then
'                num = num + 1
'            End If
'        End If
'        '次のフォルダ見に行く
'        dName = Dir()
'    Loop
'
'    DirectoryNumber = num
'
'End Function
'
'
''シートを追加するプロシージャ
'Function GetImage_AddSheets(argPath As String, argArrDirectory() As String) As String()
'
'    Dim dName As String
'    Dim i As Long
'    Dim tmpSheet As Worksheet
'
'    '選択したフォルダ名取得
'    dName = Dir(argPath & "\*.*", vbDirectory)
'
'    'フォルダ数ループ
'    Do While dName <> ""
'        'dNameがフォルダかチェック
'        If GetAttr(argPath & "\" & dName) And vbDirectory Then
'            'dNameが、現在フォルダまたは親フォルダかチェック
'            If dName <> "." And dName <> ".." Then
'                i = i + 1
'                '追加するシートと同名のシートがあるかチェック(１周目のみ実行)
'                If i = 1 Then
'                    'シート数が1個ならシート名変更(シート名&"1")
'                    If Sheets.count = 1 Then
'                        ActiveSheet.Name = ActiveSheet.Name & "1"
'                    'シート数が複数
'                    Else
'                        'シート数ループ
'                        For Each tmpSheet In Sheets
'                            'dNameがブックにあるシートと同名なら削除
'                            If tmpSheet.Name = dName Then
'                                '確認画面非表示
'                                Application.DisplayAlerts = False
'                                tmpSheet.Delete
'                                Application.DisplayAlerts = True
'                                Exit For
'                            End If
'                        Next
'                    End If
'                End If
'                'ワークシートを追加し、名前付ける
'                Worksheets.Add After:=Worksheets(Worksheets.count)
'                dName = Left(dName, 30)
'                ActiveSheet.Name = dName
'                'MsgBox ActiveSheet.CodeName
'
'                '追加したシート以外のシートを削除(１周目のみ実行)
''                If i = 1 Then
''                    'シート数ループ
''                    For Each tmpSheet In Sheets
''                        'dNameがブックにあるシートと別名なら削除
''                        If tmpSheet.Name <> dName Then
''                            '確認画面非表示
''                            Application.DisplayAlerts = False
''                            tmpSheet.Delete
''                            Application.DisplayAlerts = True
''                        End If
''                    Next
''                End If
'
'                ' シートの色を赤色にする
'                With ActiveWorkbook.Sheets(dName).Tab
'                    .Color = 255
'                    .TintAndShade = 0
'                End With
'
'                'フォルダ名(シート名)を配列に格納
'                argArrDirectory(i - 1) = dName
'            End If
'        End If
'        '次のフォルダを見に行く
'        dName = Dir()
'    Loop
'
'    AddSheets = argArrDirectory
'
'End Function
'
'
''画像を挿入するプロシージャ
'Sub GetImage_InsertPicture(dirPath As String, FolderName As String)
'
'    Dim i As Long
'    Dim fName As String
'    Dim cellRange, pict As Object
'    Dim x, y, cellHeight, pictHeight As Long
'
'      'xとy(画像を挿入するセル)の初期値
'      x = 2
'      y = 2
'
'      'シート選択
'      Sheets(FolderName).Select
'
'      ' シートの枠線非表示
'      ActiveWindow.DisplayGridlines = False
'
'      'フォルダ内のファイル名取得
'      fName = Dir(dirPath & "\*", vbNormal)
'
'      'ファイル数ループ
'      Do While fName <> ""
'          'ファイルが画像ファイルかチェック
'          If Right(fName, 4) = ".jpg" Or _
'             Right(fName, 4) = ".png" Or _
'             Right(fName, 4) = ".bmp" Or _
'             Right(fName, 4) = ".png" Or _
'             Right(fName, 4) = ".gif" Then
'              'シートに画像を貼り付け
'              ActiveSheet.Pictures.Insert(argPath & "\" & argArrDirectory(i) & "\" & fName).Select
'              '画像の位置
'              Set pict = ActiveSheet.Shapes(Selection.Name)
'              With pict
'                .Left = Cells(y, x).Left
'                .Top = Cells(y, x).Top
'              End With
'
'              '画像に枠線追加
'              ActiveSheet.Shapes(Selection.Name).Select
'              With Selection.ShapeRange.line
'                .Visible = msoTrue
'                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
'                .ForeColor.TintAndShade = 0
''                  .ForeColor.Brightness = -0.5
'                .Transparency = 0
'              End With
'
'              ' シートの色を戻す
''              With ActiveWorkbook.Sheets(argArrDirectory(i)).Tab
''                  .Color = xlAutomatic
''                  .TintAndShade = 0
''              End With
'
'              ' 画像ファイル名を設定
'              Cells(y - 1, 1).Value = fName
'
'              '1セルの高さ取得
'              cellHeight = ActiveCell.Height
'              '画像の縦の長さ取得
'              pictHeight = pict.Height
'
'              '画像挿入位置をずらす
'              y = y + Int(pictHeight / cellHeight) + 2
'          End If
'          '次のファイルを見に行く
'          fName = Dir()
'      Loop
'      'セルA1をアクティブに
'      Range("A1").Select
'
'End Sub
'
'
'
'
