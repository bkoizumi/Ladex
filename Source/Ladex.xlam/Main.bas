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
  On Error GoTo catchError
  runFlg = True
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  ThisWorkbook.Activate
  endLine = LadexSh_Config.Cells(Rows.count, 7).End(xlUp).Row

  For line = 3 To endLine
    RegistryKey = LadexSh_Config.Range("G" & line)
    RegistrySubKey = LadexSh_Config.Range("H" & line)
    val = LadexSh_Config.Range("I" & line)
    
    If Library.getRegistry(RegistryKey, RegistrySubKey, "String") = "" Then
      Call Library.setRegistry(RegistryKey, RegistrySubKey, val)
    End If
  Next
  
  '独自関数設定----------------------------------
  Call Ctl_Hollyday.InitializeHollyday
  Call Ctl_UsrFunction.InitializeUsrFunction
  
  'ショートカットキー設定------------------------
  Call Main.setShortcutKey


  '処理終了--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end1")
  'Call init.unsetting
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
  Dim ShortcutKey As String, ShortcutFunc As String, ShortcutKey1 As String
  Const funcName As String = "Main.setShortcutKey"
  
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
  
  Call Application.OnKey("{F1}", "")
  
  endLine = LadexSh_Function.Cells(Rows.count, 1).End(xlUp).Row
  For line = 2 To endLine
    If LadexSh_Function.Range("E" & line) <> "" Then
      ShortcutKey = ""
      ShortcutKey1 = ""
      
      For Each keyVal In Split(LadexSh_Function.Range("E" & line), "+")
        If keyVal = "Ctrl" Then
          ShortcutKey = "^"
        ElseIf keyVal = "Alt" Then
          ShortcutKey = ShortcutKey & "%"
        ElseIf keyVal = "Shift" Then
          ShortcutKey = ShortcutKey & "+"
        Else
          Select Case keyVal
            Case 0 To 9
              ShortcutKey1 = ShortcutKey & "{" & 96 + keyVal & "}"
              ShortcutKey = ShortcutKey & keyVal
              
            Case Else
              ShortcutKey = ShortcutKey & keyVal
          End Select
        End If
      Next
      ShortcutFunc = "Menu.ladex_" & LadexSh_Function.Range("C" & line)
      Call Library.showDebugForm("ShortcutKey", ShortcutKey, "debug")
      Call Library.showDebugForm("Function   ", ShortcutFunc, "debug")
      
      Call Application.OnKey(ShortcutKey, ShortcutFunc)
      
      If ShortcutKey1 <> "" Then
        Call Library.showDebugForm("ShortcutKey", ShortcutKey1, "debug")
        Call Library.showDebugForm("Function   ", ShortcutFunc, "debug")
        Call Application.OnKey(ShortcutKey1, ShortcutFunc)
      End If
    End If
  Next
  
'  Call Application.OnKey("{F1}", "Ctl_Option.showVersion")


  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
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
  
  LadexSh_Style.copy
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
  targetBook.Sheets("Style").Columns("A:J").copy ThisWorkbook.Worksheets("Style").Range("A1")
  targetBook.Close
  
  Call FSO.DeleteFile(TempName, True)
  
  Call Ctl_Style.スタイル削除
  Call Library.endScript
End Function

'==================================================================================================
Function 右クリックメニュー(Target As Range, Cancel As Boolean)
  Dim menu01 As CommandBarControl
  
  Call init.setting
  
  '標準状態にリセット
  Application.CommandBars("Cell").Reset
  For Each menu01 In Application.CommandBars("Cell").Controls
    'Call Library.showDebugForm("右クリック", menu01.Caption, "debug")
    
    If menu01.Caption Like "*[複合表として 追加操作]*" Then
      menu01.Visible = False
    End If
  Next

  
  With Application.CommandBars("Cell").Controls.add(Before:=1, Type:=msoControlPopup, Temporary:=True)
    .Caption = thisAppName
    If Not (Target.count = Rows.count Or Target.count = Columns.count) Then
      With .Controls.add(Temporary:=True)
        .Caption = "行列を入れ替えて貼付け"
        .OnAction = "menu.ladex_行例を入れ替えて貼付け"
      End With
    End If
    With .Controls.add(Temporary:=True)
      .BeginGroup = True
      .Caption = "行の挿入"
      .OnAction = "menu.ladex_行挿入"
    End With
    With .Controls.add(Temporary:=True)
      .Caption = "列の挿入"
      .OnAction = "menu.ladex_列挿入"
    End With
  End With
  


  Application.CommandBars("Cell").ShowPopup
  Application.CommandBars("Cell").Reset
  Cancel = True
End Function









