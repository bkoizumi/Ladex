'-------------------------------------------------------------------------------
' Excel�t�@�C���̉E�N���b�N�u�ی�r���[�ŊJ���v��L���ɂ���X�N���v�g
'
Option Explicit

On Error Resume Next

If WScript.Arguments.Count = 0 Then

    '�������g���Ǘ��Ҍ����Ŏ��s
    With CreateObject("Shell.Application")
        .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ dummy", "", "runas"
    End With

    WScript.Quit

End If


With WScript.CreateObject("WScript.Shell")

    '�V�t�g�������Ȃ��Ă����j���[���\�������悤�ɂ���悤�ɁuExtended�v�L�[���폜
    .RegDelete "HKCR\Excel.Sheet.8\shell\ViewProtected\Extended"
    .RegDelete "HKCR\Excel.Sheet.12\shell\ViewProtected\Extended"
    .RegDelete "HKCR\Excel.SheetMacroEnabled.12\shell\ViewProtected\Extended"

    Err.Clear

    '�ǂݎ���p��L���ɂ���
    .RegWrite "HKCR\Excel.Sheet.8\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.Sheet.12\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.SheetMacroEnabled.12\shell\ViewProtected\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"

End With

If Err.Number = 0 Then
    MsgBox "����Ɋ֘A�t����ύX���܂����B", vbInformation + vbOkOnly, "�ی�r���[�L����"
Else
    MsgBox "�G���[���������܂����B", vbCritical + vbOkOnly, "�ی�r���[�L����"
End IF
