'-------------------------------------------------------------------------------
' Excelファイルの右クリック「保護ビューで開く」を有効にするスクリプト
'
Option Explicit

On Error Resume Next

If WScript.Arguments.Count = 0 Then

    '自分自身を管理者権限で実行
    With CreateObject("Shell.Application")
        .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ dummy", "", "runas"
    End With

    WScript.Quit

End If


With WScript.CreateObject("WScript.Shell")

    'シフトを押さなくてもメニューが表示されるようにするように「Extended」キーを削除
    .RegDelete "HKCR\Excel.Sheet.8\shell\ViewProtected\Extended"
    .RegDelete "HKCR\Excel.Sheet.12\shell\ViewProtected\Extended"
    .RegDelete "HKCR\Excel.SheetMacroEnabled.12\shell\ViewProtected\Extended"

    Err.Clear

    '読み取り専用を有効にする
    .RegWrite "HKCR\Excel.Sheet.8\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.Sheet.12\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.SheetMacroEnabled.12\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"

End With

If Err.Number = 0 Then
    MsgBox "正常に関連付けを変更しました。", vbInformation + vbOkOnly, "保護ビュー有効化"
Else
    MsgBox "エラーが発生しました。", vbCritical + vbOkOnly, "保護ビュー有効化"
End IF
