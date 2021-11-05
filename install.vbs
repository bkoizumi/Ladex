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

Old_addInName = "Liadex"
Old_addInFileName = "Liadex.xlam"

IF MsgBox(addInName & " をインストールしますか？", vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF


Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
    MsgBox "Zipファイルを展開してから実行してください。", vbExclamation, addInName
    WScript.Quit
END IF

UnIstallPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & Old_addInFileName

If objFileSys.FileExists(UnIstallPath) = True Then
  'Excel インスタンス化
  Set objExcel = CreateObject("Excel.Application")
  objExcel.Workbooks.Add
  'アドイン登録解除
  For i = 1 To objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = Old_addInFileName Then
      objAddin.Installed = False
    End If
  Next
  objExcel.Quit

  'ファイル削除
  objFileSys.DeleteFile UnIstallPath , True
  IF Err.Number = 0 THEN
  ELSE
    MsgBox "アンインストールに失敗しました" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, Old_addInName
    WScript.Quit
  End IF
End If



'インストール------------------
With CreateObject("Excel.Application")

  'インストール先パスの作成
  strPath = .UserLibraryPath
  installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

  imageFolder = objWshShell.SpecialFolders("Appdata") & "\Bkoizumi\Ladex\"

  'インストールフォルダがない場合は作成
  IF Not objFileSys.FolderExists(strPath) THEN
      objFileSys.CreateFolder(strPath)
  END IF


  'ファイルコピー(上書き)
  objFileSys.CopyFile  addInFileName ,installPath , True

  'イメージフォルダがない場合は作成
  IF Not objFileSys.FolderExists(imageFolder) THEN
      objFileSys.CreateFolder(imageFolder)
  END IF

  'イメージフォルダをコピー(上書き)
  objFileSys.CopyFolder  "Source\Ladex\*" ,imageFolder , True

  'アドイン登録
  .Workbooks.Add
  Set objAddin = .AddIns.Add(installPath, True)
  objAddin.Installed = True

  'Excel 終了
  .Quit

End WIth

IF Err.Number = 0 THEN
    'ファイルのプロパティ表示
    MsgBox "インターネットから取得したファイルはExcelよりブロックされる場合があります。" & vbCrlf & "プロパティウィンドウを開きますので「ブロックの解除」を行ってください。" & vbCrLf & vbCrLf & "プロパティに「ブロックの解除」が表示されない場合は特に操作の必要はありません。", vbExclamation, addInName
    CreateObject("Shell.Application").NameSpace(strPath).ParseName(addInFileName).InvokeVerb("properties")

    MsgBox "アドインのインストールが終了しました。", vbInformation, addInName
ELSE
    MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName
    WScript.Quit

End IF

Set objFileSys = Nothing
Set objWshShell = Nothing
Set objAddin =  Nothing
