' -------------------------------------------------------------------------------
' Addin �C���X�g�[���X�N���v�g Ver.1.0.0
' -------------------------------------------------------------------------------
' �Q�l�T�C�g
' ���� SE �̂Ԃ₫
' VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@
' https://www.aruse.net/entry/2018/09/13/081734

' RelaxTools Addin for Excel 2013/2016/2019/Office365(Desktop)
' https://software.opensquare.net/relaxtools/
' -------------------------------------------------------------------------------
Option Explicit
On Error Resume Next

Dim installPath
Dim UnIstallPath
Dim addInName
Dim addInFileName
Dim Old_addInName
Dim Old_addInFileName
Dim objExcel
Dim objAddin
Dim imageFolder
Dim appFile
Dim objWshShell
Dim objFileSys
Dim strPath
Dim objFolder
Dim objFile
dim i

'�A�h�C������ݒ�
addInName = "Ladex"
addInFileName = "Ladex.xlam"

Old_addInName = "Liadex"
Old_addInFileName = "Liadex.xlam"

IF MsgBox(addInName & " ���C���X�g�[�����܂����H", vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF


Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
    MsgBox "Zip�t�@�C����W�J���Ă�����s���Ă��������B", vbExclamation, addInName
    WScript.Quit
END IF

UnIstallPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & Old_addInFileName

If objFileSys.FileExists(UnIstallPath) = True Then
  'Excel �C���X�^���X��
  Set objExcel = CreateObject("Excel.Application")
  objExcel.Workbooks.Add
  '�A�h�C���o�^����
  For i = 1 To objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = Old_addInFileName Then
      objAddin.Installed = False
    End If
  Next
  objExcel.Quit

  '�t�@�C���폜
  objFileSys.DeleteFile UnIstallPath , True
  IF Err.Number = 0 THEN
  ELSE
    MsgBox "�A���C���X�g�[���Ɏ��s���܂���" & vbCrLF & "Excel���N�����Ă���ꍇ�͏I�����Ă��������B", vbExclamation, Old_addInName
    WScript.Quit
  End IF
End If



'�C���X�g�[��------------------
With CreateObject("Excel.Application")

  '�C���X�g�[����p�X�̍쐬
  strPath = .UserLibraryPath
  installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

  imageFolder = objWshShell.SpecialFolders("Appdata") & "\Bkoizumi\Ladex\"

  '�C���X�g�[���t�H���_���Ȃ��ꍇ�͍쐬
  IF Not objFileSys.FolderExists(strPath) THEN
      objFileSys.CreateFolder(strPath)
  END IF


  '�t�@�C���R�s�[(�㏑��)
  objFileSys.CopyFile  addInFileName ,installPath , True

  '�C���[�W�t�H���_���Ȃ��ꍇ�͍쐬
  IF Not objFileSys.FolderExists(imageFolder) THEN
      objFileSys.CreateFolder(imageFolder)
  END IF

  '�C���[�W�t�H���_���R�s�[(�㏑��)
  objFileSys.CopyFolder  "Source\Ladex\*" ,imageFolder , True

  '�A�h�C���o�^
  .Workbooks.Add
  Set objAddin = .AddIns.Add(installPath, True)
  objAddin.Installed = True

  'Excel �I��
  .Quit

End WIth

IF Err.Number = 0 THEN
    '�t�@�C���̃v���p�e�B�\��
    MsgBox "�C���^�[�l�b�g����擾�����t�@�C����Excel���u���b�N�����ꍇ������܂��B" & vbCrlf & "�v���p�e�B�E�B���h�E���J���܂��̂Łu�u���b�N�̉����v���s���Ă��������B" & vbCrLf & vbCrLf & "�v���p�e�B�Ɂu�u���b�N�̉����v���\������Ȃ��ꍇ�͓��ɑ���̕K�v�͂���܂���B", vbExclamation, addInName
    CreateObject("Shell.Application").NameSpace(strPath).ParseName(addInFileName).InvokeVerb("properties")

    MsgBox "�A�h�C���̃C���X�g�[�����I�����܂����B", vbInformation, addInName
ELSE
    MsgBox "�G���[���������܂����B" & vbCrLF & "Excel���N�����Ă���ꍇ�͏I�����Ă��������B", vbExclamation, addInName
    WScript.Quit

End IF

Set objFileSys = Nothing
Set objWshShell = Nothing
Set objAddin =  Nothing
