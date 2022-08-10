Attribute VB_Name = "Menu"
Option Explicit

'**************************************************************************************************
' * ショートカットキーからの呼び出し用
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub ladex_Notation_R1C1()
  Call Ctl_Sheet.R1C1表記
End Sub

Sub ladex_resetStyle()
  Call Ctl_Style.スタイル初期化
End Sub

Sub ladex_delStyle()
  Call Ctl_Style.スタイル削除
End Sub

Sub ladex_setStyle()
  Call Ctl_Style.スタイル設定
End Sub

Sub ladex_del_CellNames()
  Call Ctl_Book.名前定義削除
End Sub

Sub ladex_disp_SVGA12()
  Call Ctl_Window.画面サイズ変更(612, 432)
End Sub

Sub ladex_disp_HD15_6()
  Call Ctl_Window.画面サイズ変更(1920, 1080)
End Sub

Sub ladex_シート一覧取得()
  Call Ctl_Book.シートリスト取得
End Sub

Sub ladex_セル選択()
  Application.Goto Reference:=Range("A1"), Scroll:=True
End Sub

Sub ladex_セル選択_保存()
  Application.Goto Reference:=Range("A1"), Scroll:=True
  ActiveWorkbook.Save
End Sub

Sub ladex_全セル表示()
  Call Ctl_Sheet.すべて表示
End Sub

Sub ladex_セルとシート選択()
  Call Ctl_Sheet.A1セル選択
End Sub

Sub ladex_セルとシート_保存()
  Call Ctl_Sheet.A1セル選択
  ActiveWorkbook.Save
End Sub

Sub ladex_標準画面()
  Call Ctl_Sheet.標準画面
End Sub

Sub ladex_画面最大化()
  Call Ctl_Zoom.Zoom01
End Sub

Sub ladex_初期表示倍率()
  Call Ctl_Zoom.defaultZoom
End Sub


Sub ladex_セル調整_幅()
  Call Ctl_Cells.セル幅調整
End Sub

Sub ladex_セル調整_高さ()
  Call Ctl_Cells.セル高さ調整
End Sub

Sub ladex_セル調整_両方()
  Call Ctl_Cells.セル幅調整
  Call Ctl_Cells.セル高さ調整
End Sub

Sub ladex_セル幅取得()
  Call Library.getColumnWidth
End Sub

Sub ladex_シート管理_フォーム表示(Optional control As IRibbonControl)
  Call Ctl_Sheet.シート管理_フォーム表示
End Sub

Sub ladex_文頭文末のスペース削除()
  Call Ctl_Cells.Trim01
End Sub

Sub ladex_中黒点付与()
  Call Ctl_Cells.中黒点付与
End Sub

Sub ladex_連番追加()
  Call Ctl_Cells.連番追加
End Sub

Sub ladex_全半角変換()
  Call Ctl_Cells.英数字全半角変換
End Sub

Sub ladex_取り消し線()
  Call Ctl_Cells.取り消し線設定
End Sub

Sub ladex_コメント挿入()
  Call init.unsetting
  Call Ctl_Cells.コメント挿入
End Sub

Sub ladex_コメント削除()
  Call Ctl_Cells.コメント削除
End Sub

Sub ladex_行挿入()
  Call Ctl_Cells.行挿入
End Sub

Sub ladex_列挿入()
  Call Ctl_Cells.列挿入
End Sub


Sub ladex_コメント整形()
  Call Ctl_format.コメント整形
End Sub

Sub ladex_行例を入れ替えて貼付け()
  Call Ctl_Cells.行例を入れ替えて貼付け
End Sub

Sub ladex_数式エラー防止()
  Call Ctl_Formula.エラー防止
End Sub

Sub ladex_整形_1()
  Call Ctl_format.移動やサイズ変更をする
End Sub

Sub ladex_整形_2()
  Call Ctl_format.移動する
End Sub

Sub ladex_整形_3()
  Call Ctl_format.移動やサイズ変更をしない
End Sub

Sub ladex_上下の余白ゼロ()
  Call Ctl_format.上下余白ゼロ
End Sub

Sub ladex_画像保存()
  Call Ctl_Image.saveSelectArea2Image
End Sub

Sub ladex_罫線_クリア()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_クリア
  Call Library.endScript
End Sub

Sub ladex_罫線_クリア_中央線_横()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_中央線削除_横
  Call Library.endScript
End Sub

Sub ladex_罫線_クリア_中央線_縦()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_中央線削除_縦
  Call Library.endScript
End Sub

Sub ladex_罫線_表_実線()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_格子
  Call Library.endScript
End Sub

Sub ladex_罫線_表_破線A()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_表
  Call Library.endScript
End Sub

Sub ladex_罫線_表_破線B()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_格子
  Call Library.罫線_実線_水平
  Call Library.罫線_実線_囲み
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_水平()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_水平
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_垂直()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_垂直
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_左()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_左
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_右()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_右
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_左右()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_左右
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_上()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_上
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_下()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_下
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_上下()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_上下
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_囲み()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_囲み
  Call Library.endScript
End Sub

Sub ladex_罫線_破線_格子()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_破線_格子
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_水平()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_水平
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_垂直()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_垂直
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_左右()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_左右
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_上下()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_上下
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_囲み()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_囲み
  Call Library.endScript
End Sub

Sub ladex_罫線_実線_格子()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_実線_格子
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_左()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_左
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_左右()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_左右
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_上()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_上
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_下()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_下
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_上下()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_上下
  Call Library.endScript
End Sub

Sub ladex_罫線_二重線_囲み()
  Call init.setting
  Call Library.startScript
  Call Library.罫線_二重線_囲み
  Call Library.endScript
End Sub

Sub ladex_連番設定()
  Call Ctl_Cells.連番設定
End Sub

Sub ladex_連番生成()
  Call Ctl_Cells.連番追加
End Sub

Sub ladex_桁数固定数値()
  Call Ctl_sampleData.数値_桁数固定(Selection.CountLarge)
End Sub

Sub ladex_範囲指定数値()
  Call Ctl_sampleData.数値_範囲
End Sub

Sub ladex_データ生成_姓()
  Call Ctl_sampleData.名前_姓(Selection.CountLarge)
End Sub

Sub ladex_データ生成_名()
  Call Ctl_sampleData.名前_名(Selection.CountLarge)
End Sub

Sub ladex_データ生成_氏名()
  Call Ctl_sampleData.名前_フルネーム(Selection.CountLarge)
End Sub

Sub ladex_データ生成_日付()
  Call Ctl_sampleData.日付_日(Selection.CountLarge)
End Sub

Sub ladex_データ生成_時間()
  Call Ctl_sampleData.日付_時間(Selection.CountLarge)
End Sub

Sub ladex_データ生成_日時()
  Call Ctl_sampleData.日時(Selection.CountLarge)
End Sub

Sub ladex_データ生成_文字()
  Call Ctl_sampleData.その他_文字
End Sub

Sub ladex_カスタム01()
  Call Ctl_カスタム.カスタム01
End Sub
Sub ladex_カスタム02()
  Call Ctl_カスタム.カスタム02
End Sub
Sub ladex_カスタム03()
  Call Ctl_カスタム.カスタム03
End Sub
Sub ladex_カスタム04()
  Call Ctl_カスタム.カスタム04
End Sub
Sub ladex_カスタム05()
  Call Ctl_カスタム.カスタム05
End Sub


