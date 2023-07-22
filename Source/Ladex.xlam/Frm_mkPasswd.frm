VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_mkPasswd 
   Caption         =   "パスワード生成"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   OleObjectBlob   =   "Frm_mkPasswd.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_mkPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()

  Call init.setting
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Width) / 2)
  
  '文字数
  If dicVal("MKPW_PWLen") = "" Then
     PWLen.Value = 12
  Else
    PWLen.Value = CInt(dicVal("MKPW_PWLen"))
  End If
  
  
  'a-z
  If dicVal("MKPW_LowerCaseFlg") = "True" Then
    LowerCaseFlg.Value = True
  End If
  
  'A-Z
  If dicVal("MKPW_UpperCaseFlg") = "True" Then
    UpperCaseFlg.Value = True
  End If
  
  '0-9
  If dicVal("MKPW_NumberFlg") = "True" Then
    NumberFlg.Value = True
  End If
  
  '記号
  If dicVal("MKPW_SymbolVal") = "" Then
    dicVal("MKPW_SymbolVal") = "!@#$%&"
  End If
  If dicVal("MKPW_SymbolFlg") = "True" Then
    SymbolVal.Enabled = True
  Else
    SymbolVal.Enabled = False
  End If
  
  'オプション
  If dicVal("MKPW_OptionsFlg") = "True" Then
    OptionsFlg.Value = True
  End If

End Sub





'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub PWLen_Change()
  If PWLen.Value < 1 Then
    PWLen.Value = 1
  End If
End Sub


Private Sub SpinButton1_SpinUp()
  PWLen.Value = PWLen.Value + 1
End Sub

Private Sub SpinButton1_SpinDown()
  PWLen.Value = PWLen.Value - 1
End Sub

Private Sub SymbolFlg_Click()
  If SymbolFlg.Value Then
    SymbolVal.Enabled = True
  Else
    SymbolVal.Enabled = False
  End If
  
End Sub


'キャンセル処理 -----------------------------------------------------------------------------------
Private Sub Cancel_Click()
  Call Library.setRegistry("Main", "MKPW_LowerCaseFlg", LowerCaseFlg.Value)
  Call Library.setRegistry("Main", "MKPW_UpperCaseFlg", UpperCaseFlg.Value)
  Call Library.setRegistry("Main", "MKPW_NumberFlg", NumberFlg.Value)
  Call Library.setRegistry("Main", "MKPW_SymbolFlg", SymbolFlg.Value)
  Call Library.setRegistry("Main", "MKPW_SymbolVal", SymbolVal.Value)
  Call Library.setRegistry("Main", "MKPW_OptionsFlg", OptionsFlg.Value)
  Call Library.setRegistry("Main", "MKPW_PWLen", PWLen.Value)

  Unload Me
End Sub

'パスワード生成 -----------------------------------------------------------------------------------
Private Sub run_Click()
  Call makePassword
End Sub
  
' クリップボードにコピー --------------------------------------------------------------------------
Private Sub copy_Click()
  If passWord.Value <> "" Then
    With CreateObject("Forms.TextBox.1")
      .MultiLine = True
      .Text = passWord.Value
      .SelStart = 0
      .SelLength = .TextLength
      .copy
    End With
  End If
End Sub


'**************************************************************************************************
' * ランダムなパスワード生成関数
' *
' * @Link   https://thom.hateblo.jp/entry/2017/11/29/213607
'**************************************************************************************************
'パスワード生成 -----------------------------------------------------------------------------------
Private Function makePassword()
  Dim passWordVal As String, val As String
  Dim i As Integer, n
  
  Call init.setting
  passWordVal = ""
  
  Do While Len(passWordVal) <= CInt(PWLen.Value)
    'a-z
    If LowerCaseFlg.Value Then
      passWordVal = passWordVal & RandomCharPicker(LCase(HalfWidthCharacters), passWordVal)
    End If
    
    'A-Z
    If UpperCaseFlg.Value Then
      passWordVal = passWordVal & RandomCharPicker(HalfWidthCharacters, passWordVal)
    End If
  
    '0-9
    If NumberFlg.Value Then
      passWordVal = passWordVal & RandomCharPicker(HalfWidthDigit, passWordVal)
    End If
  
    '記号
    If SymbolFlg.Value Then
      passWordVal = passWordVal & RandomCharPicker(SymbolVal, passWordVal)
    End If
    
    'オプション
    If OptionsFlg.Value Then
      passWordVal = Replace(passWordVal, "0", "")
      passWordVal = Replace(passWordVal, "1", "")
      passWordVal = Replace(passWordVal, "l", "")
      passWordVal = Replace(passWordVal, "O", "")
    End If
  Loop
  passWordVal = Left(ShuffleString(passWordVal), CInt(PWLen.Value))
  
  Call Library.showDebugForm("passWordVal", passWordVal, "debug")
  passWord.Value = passWordVal

End Function


'ランダムな1文字を取得-----------------------------------------------------------------------------
Function RandomCharPicker(Source, Optional passWordVal As String)
  Dim location As String
  Dim pickVal As String
  Dim n As Integer, reRunCnt As Integer
  
  reRunCnt = 0
LBL_reStart:
  Randomize
  n = Int((Len(Source) - 1 + 1) * Rnd + 1)
  pickVal = Mid(Source, n, 1)
  
  If InStr(passWordVal, pickVal) = 0 Then
    RandomCharPicker = pickVal
  Else
    reRunCnt = reRunCnt + 1
    If reRunCnt <= 2 Then
      GoTo LBL_reStart
    Else
      RandomCharPicker = pickVal
    End If
  End If
  
End Function

'文字列をシャッフル--------------------------------------------------------------------------------
Function ShuffleString(Source)
  Dim c As Collection: Set c = New Collection
  Dim i As Long
  
  'まず1文字ずつコレクションに格納していく
  For i = 1 To Len(Source)
    c.add Mid(Source, i, 1)
  Next
  
  'コレクションが空になるまで、ランダムに取り出す。
  Dim ret As String
  Dim location As Long
  Do While c.count > 0
    Randomize
    location = Int((c.count - 1 + 1) * Rnd + 1)
    ret = ret & c(location)
    c.Remove location
  Loop
  ShuffleString = ret
  End Function
