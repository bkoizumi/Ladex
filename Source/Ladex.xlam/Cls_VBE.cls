VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_VBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'
''プライベートフィールドの宣言
'Private WithEvents m_CBE As CommandBarEvents
'
''初期化メソッド
''このクラスのインスタンス生成後、すぐに呼んで。
''JavaやVB.Netではコンストラクタに初期化処理を書くんだけど、VBAではコンストラクタ(≒Initializeイベントプロシージャ)に引 数を持たせられないから、こういう形にした。
'Public Sub InitializeInstance(ByVal cbc As CommandBarControl)
'    Set m_CBE = Application.VBE.Events.CommandBarEvents(cbc)
'End Sub
'
''ボタンがクリックされた時の処理
'Private Sub m_CBE_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    Debug.Print "Click" 'イミディエイトウインドウに"Click"という文字列を出力
'End Sub
