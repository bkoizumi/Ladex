Attribute VB_Name = "Ctl_ImportFile"
'Option Explicit

'**************************************************************************************************
' * �t�@�C���C���|�[�g
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
  
  '�����J�n--------------------------------------
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
  
  'json�t�@�C���̓ǂݍ���-------------------------
  'Library.getFilePath(ActiveWorkbook.path, "", "Json�t�@�C����I�����Ă��������B", "json")
  
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
      Debug.Print items & "�F" & objJson(tmp).Item(items)
    
    Next
    Debug.Print tmp & "�F" & objJson(tmp)

  Next
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

