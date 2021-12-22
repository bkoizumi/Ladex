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

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")


'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
'�A�h�C���o�^����
For i = 1 To objExcel.Addins.Count
  Set objAddin = objExcel.Addins.item(i)
  If objAddin.Name = Old_addInFileName or objAddin.Name = addInFileName Then
    objAddin.Installed = False
    WScript.echo(objAddin.Name)
  End If
Next
objExcel.Quit


UnIstallPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & Old_addInFileName
If objFileSys.FileExists(UnIstallPath) = True Then
  '�t�@�C���폜
  objFileSys.DeleteFile UnIstallPath , True
  IF Err.Number = 0 THEN
  ELSE
    WScript.echo(Err.Number)
    WScript.Quit
  End IF
End If

UnIstallPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName
If objFileSys.FileExists(UnIstallPath) = True Then
  '�t�@�C���폜
  objFileSys.DeleteFile UnIstallPath , True
  IF Err.Number = 0 THEN
  ELSE
    WScript.echo(Err.Number)
    WScript.Quit
  End IF
End If



WScript.echo(Err.Number)

Set objFileSys = Nothing
Set objWshShell = Nothing
Set objAddin =  Nothing
