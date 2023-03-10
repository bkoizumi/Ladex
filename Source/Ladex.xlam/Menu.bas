Attribute VB_Name = "Menu"
Option Explicit

'**************************************************************************************************
' * 各機能呼び出し
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 各機能呼び出し(shortcutName As String)
  
  Const funcName As String = "Menu.各機能呼び出し"

  '処理開始--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")
  Call Library.showDebugForm("shortcutName", shortcutName, "debug")
  '----------------------------------------------

  Select Case shortcutName

    'お気に入り----------------------------------
    Case "Favorite_detail": Call Ctl_Favorite.詳細表示
    Case "追加": Call Ctl_Favorite.追加

    'Option--------------------------------------
    Case "showOption__": Call Ctl_Option.showOption
    Case "スタイル出力": Call Ctl_Style.スタイル出力
    Case "スタイル取込": Call Ctl_Style.スタイル取込
    Case "showVersion_": Call Ctl_Option.showVersion
    Case "showHelp____": Call Ctl_Option.showHelp
    Case "Addin解除___": Call Ctl_Option.Addin解除
    Case "初期化______": Call Ctl_Option.初期化
    
    'ブック管理----------------------------------
    Case "スタイル初期化___": Call Ctl_Style.スタイル初期化
    Case "スタイル削除_____": Call Ctl_Style.スタイル削除
    Case "スタイル設定_____": Call Ctl_Style.スタイル設定
    Case "名前定義削除_____": Call Ctl_Book.名前定義削除
    Case "B_A1セル選択_____": Call Ctl_Book.A1セル選択
    Case "B_A1セル選択_保存": Call Ctl_Book.A1セル選択_保存
    Case "印刷範囲枠表示___": Call Ctl_Book.印刷範囲表示
    Case "印刷範囲枠非表示_": Call Ctl_Book.印刷範囲非表示
    Case "シート一覧取得___": Call Ctl_Book.シート一覧取得
    Case "連続シート追加___": Call Ctl_Book.連続シート追加
    Case "指定倍率_________": Call Ctl_Zoom.指定倍率
    Case "標準画面_________": Call Ctl_Book.標準画面
    Case "シート管理_______": Call Ctl_Book.シート管理

    'シート管理----------------------------------
    Case "S_A1セル選択______": Call Ctl_Sheet.A1セル選択
    Case "S_A1セル選択_保存_": Call Ctl_Sheet.A1セル選択_保存
    Case "すべて表示________": Call Ctl_Sheet.すべて表示
    Case "不要データ削除____": Call Ctl_Sheet.不要データ削除
    Case "体裁一括変更______": Call Ctl_Sheet.体裁一括変更
    Case "フォント一括変更__": Call Ctl_Sheet.フォント一括変更

    'その他管理--------------------------------
    Case "ファイル情報取得": Call Ctl_File.ファイル情報取得
    Case "フォルダ生成____": Call Ctl_File.フォルダ生成
    Case "画像貼付け______": Call Ctl_File.画像貼付け
    Case "はんこ_確認印___": Call Ctl_Stamp.はんこ_確認印
    Case "はんこ_名前_____": Call Ctl_Stamp.はんこ_名前
    Case "はんこ_済印_____": Call Ctl_Stamp.はんこ_済印
    Case "QRコード生成____": Call Ctl_shap.QRコード生成
    Case "パスワード生成__": Call Ctl_shap.パスワード生成
    Case "カスタム関数01__": Call Ctl_カスタム.カスタム関数01
    Case "カスタム関数02__": Call Ctl_カスタム.カスタム関数02
    Case "カスタム関数03__": Call Ctl_カスタム.カスタム関数03
    Case "カスタム関数04__": Call Ctl_カスタム.カスタム関数04
    Case "カスタム関数05__": Call Ctl_カスタム.カスタム関数05


    Case "R1C1表記": Call Ctl_Book.R1C1表記

    'ズーム--------------------------------------
    Case "全画面表示": Call Ctl_Zoom.全画面表示



    'セル調整------------------------------------
    Case "セル自動調整_幅____": Call Ctl_Cells.セル自動調整_幅
    Case "セル自動調整_高さ__": Call Ctl_Cells.セル自動調整_高さ
    Case "セル自動調整_両方__": Call Ctl_Cells.セル自動調整_両方
    Case "セル固定設定_幅____": Call Ctl_Cells.セル固定設定_幅
    Case "セル固定設定_高さ__": Call Ctl_Cells.セル固定設定_高さ
    Case "セル固定設定_両方__": Call Ctl_Cells.セル固定設定_両方
    Case "セル固定設定_高さ15": Call Ctl_Cells.セル固定設定_高さ15
    Case "セル固定設定_高さ30": Call Ctl_Cells.セル固定設定_高さ30

    'セル編集------------------------------------
    Case "削除_前後のスペース_________": Call Ctl_Cells.削除_前後のスペース
    Case "削除_全スペース_____________": Call Ctl_Cells.削除_全スペース
    Case "削除_改行___________________": Call Ctl_Cells.削除_改行
    Case "削除_定数___________________": Call Ctl_Cells.削除_定数
    Case "削除_コメント_______________": Call Ctl_Cells.削除_コメント
    Case "追加_文頭に中黒点___________": Call Ctl_Cells.追加_文頭に中黒点
    Case "追加_文頭に連番_____________": Call Ctl_Cells.追加_文頭に連番
    Case "追加_コメント_______________": Call Ctl_Cells.追加_コメント
    Case "上書_文頭に連番_____________": Call Ctl_Cells.上書_文頭に連番
    Case "上書_ゼロ___________________": Call Ctl_Cells.上書_ゼロ
    Case "変換_全半角_________________": Call Ctl_Cells.変換_全角⇒半角
    Case "変換_半全角_________________": Call Ctl_Cells.変換_半角⇒全角
    Case "変換_大小___________________": Call Ctl_Cells.変換_大文字⇒小文字
    Case "変換_小大___________________": Call Ctl_Cells.変換_小文字⇒大文字
    Case "変換_URLエンコード__________": Call Ctl_Cells.変換_URLエンコード
    Case "変換_URLデコード____________": Call Ctl_Cells.変換_URLデコード
    Case "変換_Unicodeエスケープ______": Call Ctl_Cells.変換_Unicodeエスケープ
    Case "変換_Unicodeアンエスケープ__": Call Ctl_Cells.変換_Unicodeアンエスケープ
    Case "変換_Base64エンコード_______": Call Ctl_Cells.変換_Base64エンコード
    Case "変換_Base64デコード_________": Call Ctl_Cells.変換_Base64デコード
    Case "変換_数値を丸数字___________": Call Ctl_Cells.変換_数値⇒丸数字
    Case "変換_丸数字を数値___________": Call Ctl_Cells.変換_丸数字⇒数値
    Case "設定_取り消し線_____________": Call Ctl_Cells.設定_取り消し線
    Case "貼付_行例入れ替え__________": Call Ctl_Cells.貼付_行例入れ替え

    '数式編集------------------------------------
    Case "数式確認________________": Call Ctl_Formula.数式確認
    Case "数式追加_エラー防止_空白": Call Ctl_Formula.エラー防止_空白
    Case "数式追加_エラー防止_ゼロ": Call Ctl_Formula.エラー防止_ゼロ
    Case "数式追加_ゼロ非表示_____": Call Ctl_Formula.ゼロ非表示
    Case "数式設定_行番号_________": Call Ctl_Formula.数式設定_行番号
    Case "数式設定_シート名_______": Call Ctl_Formula.数式設定_シート名

    '整形------------------------------------
    Case "移動やサイズ変更をする____": Call Ctl_format.移動やサイズ変更をする
    Case "移動する__________________": Call Ctl_format.移動する
    Case "移動やサイズ変更をしない__": Call Ctl_format.移動やサイズ変更をしない
    Case "上下の余白をゼロにする____": Call Ctl_format.上下余白ゼロ
    Case "左右の余白をゼロにする____": Call Ctl_format.左右余白ゼロ
    Case "文字サイズをぴったりにする": Call Ctl_format.文字サイズをぴったりにする
    Case "セル内の中央に配置________": Call Ctl_format.セル内の中央に配置

    '画像保存------------------------------------
    Case "画像保存": Call Ctl_Image.画像保存

    '罫線[クリア]--------------------------------
    Case "罫線_クリア__________": Call Library.罫線_クリア
    Case "罫線_クリア_中央線_横": Call Library.罫線_クリア_中央線_横
    Case "罫線_クリア_中央線_縦": Call Library.罫線_クリア_中央線_縦

    '罫線[表]------------------------------------
    Case "罫線_表_実線_": Call Ctl_Line.罫線_表_実線
    Case "罫線_表_破線A": Call Ctl_Line.罫線_表_破線A
    Case "罫線_表_破線B": Call Ctl_Line.罫線_表_破線B

    '罫線[破線]----------------------------------
    Case "罫線_破線_水平": Call Library.罫線_破線_水平
    Case "罫線_破線_垂直": Call Library.罫線_破線_垂直
    Case "罫線_破線_左__": Call Library.罫線_破線_左
    Case "罫線_破線_右__": Call Library.罫線_破線_右
    Case "罫線_破線_左右": Call Library.罫線_破線_左右
    Case "罫線_破線_上__": Call Library.罫線_破線_上
    Case "罫線_破線_下__": Call Library.罫線_破線_下
    Case "罫線_破線_上下": Call Library.罫線_破線_上下
    Case "罫線_破線_囲み": Call Library.罫線_破線_囲み
    Case "罫線_破線_格子": Call Library.罫線_破線_格子

    '罫線[実線]----------------------------------
    Case "罫線_実線_水平": Call Library.罫線_実線_水平
    Case "罫線_実線_垂直": Call Library.罫線_実線_垂直
    Case "罫線_実線_左右": Call Library.罫線_実線_左右
    Case "罫線_実線_上下": Call Library.罫線_実線_上下
    Case "罫線_実線_囲み": Call Library.罫線_実線_囲み
    Case "罫線_実線_格子": Call Library.罫線_実線_格子

    '罫線[二重線]----------------------------------
    Case "罫線_二重線_左__": Call Library.罫線_二重線_左
    Case "罫線_二重線_右__": Call Library.罫線_二重線_右
    Case "罫線_二重線_左右": Call Library.罫線_二重線_左右
    Case "罫線_二重線_上__": Call Library.罫線_二重線_上
    Case "罫線_二重線_下__": Call Library.罫線_二重線_下
    Case "罫線_二重線_上下": Call Library.罫線_二重線_上下
    Case "罫線_二重線_囲み": Call Library.罫線_二重線_囲み

    'データ生成-----------------------------------
    Case "データ生成_連番上書____": Call Ctl_Cells.上書_文頭に連番
    Case "データ生成_連番追加____": Call Ctl_Cells.追加_文頭に連番
    Case "データ生成_桁数固定数値": Call Ctl_sampleData.数値_桁数固定
    Case "データ生成_範囲指定数値": Call Ctl_sampleData.数値_範囲指定
    Case "データ生成_姓__________": Call Ctl_sampleData.データ生成_姓
    Case "データ生成_名__________": Call Ctl_sampleData.データ生成_名
    Case "データ生成_氏名________": Call Ctl_sampleData.データ生成_氏名
    Case "データ生成_日付________": Call Ctl_sampleData.データ生成_日付
    Case "データ生成_時間________": Call Ctl_sampleData.データ生成_時間
    Case "データ生成_日時________": Call Ctl_sampleData.データ生成_日時
    Case "データ生成_文字________": Call Ctl_sampleData.データ生成_文字
    Case "データ生成_パターン選択": Call Ctl_sampleData.データ生成_パターン選択
    
    Case Else
      Call Library.showDebugForm("リボンメニューなし", shortcutName, "Error")
      Call Library.showNotice(406, "リボンメニューなし：" & shortcutName, True)
  End Select


  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  'エラー発生時------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
