Attribute VB_Name = "Ctl_ImportFile"
'Option Explicit

'**************************************************************************************************
' * ファイルインポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Json()
  Dim wb As Workbook
  Dim tmp, items
  Dim jsonData As String
  Dim objJsons As Object, objJsonItems As Object
  
  Const funcName As String = "Ctl_ImportFile.Json"
  
  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  'jsonファイルの読み込み-------------------------
  'Library.getFilePath(ActiveWorkbook.path, "", "Jsonファイルを選択してください。", "json")
  
  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    .LoadFromFile "C:\Work\_Backup\chrome\speed-dial-2.json"
    jsonData = .ReadText
    .Close
  End With
  
  Set objJsons = JsonConverter.ParseJson(jsonData)

  Debug.Print JsonConverter.ConvertToJson(objJson, " ")
  
  For Each objJson In objJsons
    
    
    For Each objJsonItems In objJsons(objJson)
      Debug.Print items & "：" & objJson(tmp).Item(items)
    
    Next
    Debug.Print tmp & "：" & objJson(tmp)

  Next
  
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  
  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

