Attribute VB_Name = "Ctl_Ribbon"

Public ribbonUI As IRibbonUI ' リボン
Private rbButton_Visible As Boolean ' ボタンの表示／非表示
Private rbButton_Enabled As Boolean ' ボタンの有効／無効

'トグルボタン------------------------------------
Public PressT_B015 As Boolean


'**************************************************************************************************
' * リボンメニュー設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'読み込み時処理------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  
  'リボンの表示を更新する
  ribbonUI.Invalidate
  ribbonUI.ActivateTab "WBSTab"

End Function


'==================================================================================================
'トグルボタンにチェックを設定する
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
  Select Case control.ID
    Case "T_B015"
      'タイムラインに追加
      If Range(setVal("cell_Info") & ActiveCell.Row) Like "" Then
        returnedVal = True
      Else
        returnedVal = False
      End If
      
      
      
    Case Else
  End Select
End Sub


'==================================================================================================
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 2)
End Sub

'==================================================================================================
Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.ID, 3)
  Application.run setRibbonVal

End Sub


'==================================================================================================
'Supertipの動的表示
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 5)
End Sub

'==================================================================================================
Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 6)
End Sub

'==================================================================================================
Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.ID, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select


End Sub

'==================================================================================================
'Ribbonシートから内容を取得
Function getRibbonMenu(menuId As String, offsetVal As Long)

  Dim getString As String
  Dim FoundCell As Range
  Dim ribSheet As Worksheet
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Set ribSheet = ThisWorkbook.Worksheets("Ribbon")

  endLine = ribSheet.Cells(Rows.count, 1).End(xlUp).Row

  getRibbonMenu = Application.VLookup(menuId, ribSheet.Range("A2:F" & endLine), offsetVal, False)
  Call Library.endScript


  Exit Function
'エラー発生時=====================================================================================
catchError:
  getRibbonMenu = "エラー"

End Function


'**************************************************************************************************
' * リボンから実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '処理開始--------------------------------------
  runFlg = True
'  On Error GoTo catchError
  Call init.setting(True)
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  PrgP_Max = 4
  resetCellFlg = True
  '----------------------------------------------
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "CloseAddOn"
      Call Library.addonClose
    
    Case "OptionAddin解除"
      Workbooks(ThisWorkbook.Name).IsAddin = False
    
    Case "OptionAddin化"
      Workbooks(ThisWorkbook.Name).IsAddin = True
      ThisWorkbook.Save
    
    
    
    '進捗率設定----------------------------------
    Case "progress_0", "progress_25", "progress_50", "progress_75", "progress_100"
      Call Ctl_Task.進捗率設定(Replace(control.ID, "progress_", ""))
    
    
    'タスク移動----------------------------------
    Case "taskOutdent"
      Call Library.testCalAddDay
      
      Call Ctl_Task.タスク移動_左
      
    Case "taskIndent"
      Call Ctl_Task.タスク移動_右
      
    Case "taskLink"
      Call Ctl_Task.タスクのリンク設定
    Case "taskLink"
      Call Ctl_Task.タスクのリンク解除
      
    Case "chkTaskList"
      Call Ctl_Task.タスクチェック
      
      '濱さん作成VBAの呼び出し
      Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!sheet5.CommandButton7_Click"


    Case "Option"
      Call Ctl_Option.オプション画面表示
      
    Case "scrollTask"
      Call Ctl_Task.タスクにスクロール
    
    Case "addTimeLine"
      Call Ctl_Chart.タイムラインに追加(ActiveCell.Row)
        
    Case "makeChart"
         Menu.M_ガントチャート生成
         
    Case "makeCalendar"
      Call Ctl_Calendar.カレンダー生成
      'Call Ctl_Task.タスクチェック

    Case "copyProgress"
      Call Ctl_Task.進捗コピー
      
      
    Case Else
      Call Library.showDebugForm("リボンメニューなし", control.ID, "Error")
      Call Library.showNotice(406, "リボンメニューなし：" & control.ID, True)
  End Select
  
  '処理終了--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end")
  Call init.unsetting
  '----------------------------------------------
  Exit Function
  
'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


