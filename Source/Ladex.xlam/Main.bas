Attribute VB_Name = "Main"
Option Explicit

'ワークブック用変数------------------------------
'ワークシート用変数------------------------------
'グローバル変数----------------------------------


'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeBook()
  Dim RegistryKey As String, RegistrySubKey As String, val As String
  Dim line As Long, endLine As Long
  Dim regName As String
  Const funcName As String = "Main.InitializeBook"

  '処理開始--------------------------------------
  runFlg = True
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.startScript
    Call Library.showDebugForm("" & funcName, , "function")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  BK_ThisBook.Activate
  endLine = BK_sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
  
  For line = 3 To endLine
    RegistryKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistryKey") & line)
    RegistrySubKey = BK_sheetSetting.Range(BK_setVal("Cells_RegistrySubKey") & line)
    val = BK_sheetSetting.Range(BK_setVal("Cells_RegistryValue") & line)
    
    If RegistryKey <> "" Then
     Call Library.setRegistry(RegistryKey, RegistrySubKey, val)
    End If
  Next
  
  '独自関数設定----------------------------------
  Call Ctl_Hollyday.InitializeHollyday
  Call Ctl_UsrFunction.InitializeUsrFunction
  
  'ショートカットキー設定------------------------
  Call Main.setShortcutKey


  '処理終了--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * ショートカットキーの設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setShortcutKey()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim keyVal As Variant
  Dim ShortcutKey As String, ShortcutFunc As String
  Const funcName As String = "Main.setShortcutKey"
  
  '処理開始--------------------------------------
  'runFlg = True
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm("" & funcName, , "function")
    Call Library.startScript
  Else
    On Error GoTo catchError
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  endLine = BK_sheetFunction.Cells(Rows.count, 1).End(xlUp).Row
  For line = 2 To endLine
    If BK_sheetFunction.Range("B" & line) <> "" Then
      ShortcutKey = ""
      For Each keyVal In Split(BK_sheetFunction.Range("B" & line), "+")
        If keyVal = "Ctrl" Then
          ShortcutKey = "^"
        ElseIf keyVal = "Alt" Then
          ShortcutKey = ShortcutKey & "%"
        ElseIf keyVal = "Shift" Then
          ShortcutKey = ShortcutKey & "^"
        Else
          ShortcutKey = ShortcutKey & keyVal
        End If
      Next
      ShortcutFunc = "Menu.ladex_" & BK_sheetFunction.Range("D" & line)
      Call Library.showDebugForm("ShortcutKey", ShortcutKey, "debug")
      Call Library.showDebugForm("Function", ShortcutFunc, "debug")
      
      Call Application.OnKey(ShortcutKey, ShortcutFunc)
    End If
  Next
  
  'Call Application.OnKey("{F1}", "Ctl_Option.showVersion")
  Call Application.OnKey("{F1}", "")

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("  ", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("  ", , "end")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.showNotice(400, "xxxx.xxxxxxxxxx[" & Err.Number & "]" & Err.Description, True)
  Call Library.showDebugForm("[" & Err.Number & "]" & Err.Description, , "Error")
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function








'==================================================================================================
Function xxxxxxxxxx()
End Function

'**************************************************************************************************
' * 画像設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 画像設定()

  With ActiveWorkbook.ActiveSheet
    Dim AllShapes As Shapes
    Dim CurShape As Shape
    Set AllShapes = .Shapes
    
    For Each CurShape In AllShapes
      CurShape.Placement = xlMove
    Next
  End With
  
End Function



'**************************************************************************************************
' * ハイライト
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ハイライト()
'  Dim highLightFlg As String
'  Dim highLightArea As String
'
'  Call Library.startScript
'  highLightFlg = Library.getRegistry(ActiveWorkbook.Name, "HighLightFlg")
'
'  If highLightFlg = "" Then
'    Call Library.setLineColor(Selection.Address, True, Library.getRegistry("HighLightColor"))
'
'    Call Library.setRegistry(ActiveWorkbook.Name, True, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightSheet", ActiveSheet.Name, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightArea", Selection.Address, "HighLightFlg")
'
'  Else
'    highLightArea = Library.getRegistry(ActiveWorkbook.Name & "_HighLightArea")
'
'    If highLightArea = "" Then
'      highLightArea = Selection.Address
'    End If
'    Call Library.unsetLineColor(highLightArea)
'
'    Call Library.delRegistry(ActiveWorkbook.Name, "HighLightFlg")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightSheet")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightArea")
'  End If
'
'  Call Library.endScript(True)

End Function


'**************************************************************************************************
' * 設定Import / Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 設定_抽出()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting
  
  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  BK_sheetStyle.Copy
  ActiveWorkbook.SaveAs fileName:=TempName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  
  Call Library.endScript
  
  MsgBox ("修正完了後、保存し閉じてください")
End Function

'==================================================================================================
Function 設定_取込()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting

  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  Set targetBook = Workbooks.Open(TempName)
  targetBook.Sheets("Style").Columns("A:J").Copy ThisWorkbook.Worksheets("Style").Range("A1")
  targetBook.Close
  
  Call FSO.DeleteFile(TempName, True)
  
  Call Ctl_Style.スタイル削除
  Call Library.endScript
End Function











