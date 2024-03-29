' -------------------------------------------------------------------------------
' Addin インストールスクリプト Ver.1.0.0
' -------------------------------------------------------------------------------
' 参考サイト
' ある SE のつぶやき
' VBScript で Excel にアドインを自動でインストール/アンインストールする方法
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

'アドイン情報を設定
addInName = "Ladex"
addInFileName = "Ladex.xlam"

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール------------------
With CreateObject("Excel.Application")

  'インストール先パスの作成
  strPath = .UserLibraryPath
  installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

  'アドイン登録
  .Workbooks.Add
  Set objAddin = .AddIns.Add(installPath, True)
  objAddin.Installed = True

  'Excel 終了
  .Quit

End With



WScript.echo(Err.Number)


Set objFileSys = Nothing
Set objWshShell = Nothing
Set objAddin =  Nothing
