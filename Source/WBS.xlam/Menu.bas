Attribute VB_Name = "Menu"
'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting
  helpSheet.Visible = True
  helpSheet.Select
End Sub



'==================================================================================================
Sub optionKey()
Attribute optionKey.VB_ProcData.VB_Invoke_Func = "O\n14"
  Call M_オプション画面表示
End Sub
Sub centerKey()
  Call M_センター
End Sub
Sub filterKey()
  Call M_フィルター
End Sub
Sub clearFilterKey()
  Call M_すべて表示
End Sub
Sub taskCheckKey()
Attribute taskCheckKey.VB_ProcData.VB_Invoke_Func = "C\n14"
  Call M_タスクチェック
End Sub
Sub makeGanttKey()
Attribute makeGanttKey.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call M_ガントチャート生成
End Sub
Sub clearGanttKey()
Attribute clearGanttKey.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call M_ガントチャートクリア
End Sub
Sub dispAllKey()
  Call M_すべて表示
End Sub
Sub taskControlKey()
'  Call M_
End Sub
Sub ScaleKey()
'  Call M_
End Sub








Sub M_オプション画面表示()
Attribute M_オプション画面表示.VB_ProcData.VB_Invoke_Func = " \n14"
  
  Call init.setting(True)
  Call Library.startScript
  
  Call WBS_Option.オプション画面表示
  
'  Call M_カレンダー生成(True)
'  Call M_ガントチャート生成
'  Call WBS_Option.表示列設定
'
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
End Sub


Sub M_列入替え()
  Call init.setting
  
  Call Library.startScript
  Call Check.項目列チェック
  Call init.setting(True)
  
  Call Library.endScript
End Sub

Sub M_カレンダー生成(Optional flg As Boolean = False)

  Call init.setting(True)
  Call Library.startScript
  
  If flg = False Then
    Call ProgressBar.showStart
  End If
  
  '全ての行列を表示
  Cells.EntireColumn.Hidden = False
  Cells.EntireRow.Hidden = False
  
  Call Ctl_Calendar.カレンダー生成
  
  Call WBS_Option.複数の担当者行を非表示
  Call WBS_Option.表示列設定
  
  If flg = False Then
    Call ProgressBar.showEnd
  End If
  Call Library.endScript
End Sub




'**************************************************************************************************
' * 共通
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_行ハイライト()
  Call Library.startScript
  Call WBS_Option.setLineColor
  Call Library.endScript
End Sub


'--------------------------------------
Sub M_全データ削除()
  If MsgBox("データを削除します", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  
  Call Library.startScript
  Call WBS_Option.clearAll
  Call Library.endScript
End Sub


Sub M_全画面()
Attribute M_全画面.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.ScreenUpdating = False
  ActiveWindow.DisplayHeadings = False
  Application.DisplayFullScreen = True
  
  With DispFullScreenForm
    .StartUpPosition = 0
    .Top = Application.Top + 300
    .Left = Application.Left + 30
  End With
  Application.ScreenUpdating = True
  DispFullScreenForm.Show vbModeless
End Sub

Sub M_タスク操作()
Attribute M_タスク操作.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

Sub M_スケール()
Attribute M_スケール.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @Link   https://akashi-keirin.hatenablog.com/entry/2019/02/23/170807
'**************************************************************************************************
Sub M_タスクチェック()
Attribute M_タスクチェック.VB_ProcData.VB_Invoke_Func = "C\n14"
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Dim targetBookName As String
  Dim targetSheet As Object
  
  Const funcName As String = "Menu.M_タスクチェック"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  targetBookName = ActiveWorkbook.Name
  Set targetSheet = ActiveWorkbook.Worksheets("WBS")
  
  PrgP_Cnt = PrgP_Cnt + 1
  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")
  
  
  Call Ctl_Task.タスクチェック
  
  
  '濱さん作成VBAの呼び出し
  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "!sheet5.CommandButton7_Click"
  
  Set targetSheet = Nothing
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
Sub M_フィルター()
Attribute M_フィルター.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  
  With FilterForm
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
  End With
  
  FilterForm.Show
End Sub


'==================================================================================================
Sub M_すべて表示()
Attribute M_すべて表示.VB_ProcData.VB_Invoke_Func = " \n14"
  Call Library.startScript
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call WBS_Option.複数の担当者行を非表示
  Call Library.endScript
End Sub


'==================================================================================================
Sub M_進捗コピー()
  Call Task.進捗コピー
End Sub

'==================================================================================================
Sub M_タスク移動_上()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク移動_上
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク移動_下()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク移動_下
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク移動_左()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク移動_左
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク移動_右()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク移動_右
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク追加()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク追加
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク削除()

  runFlg = True
  Call init.setting
  Call Library.startScript
  Call Ctl_Task.タスク削除
  Call Library.endScript
End Sub












'==================================================================================================
Sub M_進捗率設定(progress As Long)
  Call Task.進捗率設定(progress)
End Sub

'タスクのリンク設定/解除---------------------------------------------------------------------------
Sub M_タスクのリンク設定()
  Call Ctl_Task.タスクのリンク設定
End Sub

'==================================================================================================
Sub M_タスクのリンク解除()
  Call Ctl_Task.タスクのリンク解除
End Sub

Sub M_タスクの挿入()
  Call Library.startScript
  Call init.setting
  
  Call Task.タスクの挿入
  
  Call Library.endScript
End Sub

Sub M_タスクの削除()
  Call Library.startScript
  Call init.setting
  
  Call Task.タスクの削除
  
  Call Library.endScript
End Sub

'表示モード----------------------------------------------------------------------------------------
Sub M_タスク表示_標準()
  Call Library.startScript
  
  Call init.setting
  If setVal("debugMode") <> "develop" Then
    mainSheet.Visible = True
    TeamsPlannerSheet.Visible = xlSheetVeryHidden
  End If
  
  Call init.setting(True)
  Call WBS_Option.タスク表示_標準
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  Call Library.endScript

End Sub

Sub M_タスク表示_タスク()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTask
  Call WBS_Option.setLineColor
  
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク表示_チームプランナー()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.タスク表示_チームプランナー
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
  Call Library.endScript
End Sub


'==================================================================================================
Sub M_タスクにスクロール()
  Const funcName As String = "Menu.M_タスクにスクロール"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  
  Call Task.タスクにスクロール
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub




'==================================================================================================
Sub M_タイムラインに追加()
  Const funcName As String = "Menu.M_タイムラインに追加"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  Call Ctl_Chart.タイムラインに追加(ActiveCell.Row)
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'**************************************************************************************************
' * ガントチャート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'クリア--------------------------------------------------------------------------------------------
Sub M_ガントチャートクリア()
Attribute M_ガントチャートクリア.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call Library.startScript
  Call Ctl_Chart.ガントチャート削除
  Call Library.endScript
End Sub

'生成のみ------------------------------------------------------------------------------------------
Sub M_ガントチャート生成のみ()
Attribute M_ガントチャート生成のみ.VB_ProcData.VB_Invoke_Func = "A\n14"
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("ガントチャート生成", "処理開始")
  
  Call Ctl_Chart.ガントチャート生成
  
  Call Library.showDebugForm("ガントチャート生成", "処理完了")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'==================================================================================================
'生成
Sub M_ガントチャート生成()
Attribute M_ガントチャート生成.VB_ProcData.VB_Invoke_Func = "t\n14"
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim targetBookName As String
  Dim SelectionSheet As String
  Dim objShp As Shape
  
  Const funcName As String = "Menu.M_ガントチャート生成"

  '処理開始--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 4
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  targetBookName = ActiveWorkbook.Name
  
  '濱さん作成VBAの呼び出し
  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet5.CommandButton1_Click"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.Title_Format_Check"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.CALC_MANUAL"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Chart_Module.Make_Chart"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")

  Application.run "'" & ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "'!Sheet_Module.BUTTON_CLEAR"
  Call Ctl_ProgressBar.showCount(thisAppName, PrgP_Cnt, PrgP_Max, 1, 5, "")
  
  


  
  'タイムラインに追加----------------------------
  Call Library.startScript
  Rows("6:6").RowHeight = 40
  
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "TimeLine_*" Then
      ActiveSheet.Shapes(objShp.Name).Delete
    End If
  Next
  
  
  For line = 2 To Sh_PARAM.Cells(Rows.count, 38).End(xlUp).Row
    If Sh_PARAM.Range("AL" & line).Text <> "" Then
      Call Ctl_Chart.タイムラインに追加(CLng(Sh_PARAM.Range("AL" & line).Text), True)
    End If
  Next

  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'==================================================================================================
'センター
Sub M_センター()
Attribute M_センター.VB_ProcData.VB_Invoke_Func = " \n14"

  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("センターへ移動", "処理開始")
  
  Call Ctl_Chart.センター
  
  Call Library.showDebugForm("センターへ移動", "処理完了")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excelファイル-------------------------------------------------------------------------------------
Sub M_Excelインポート()
  
  Call Library.startScript
  Call Library.showDebugForm("ファイルインポート", "処理開始")
  
  Call init.setting
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).Row
  
  If endLine > 6 Then
    If MsgBox("データを削除します", vbYesNo + vbExclamation) = vbYes Then
      Call WBS_Option.clearAll
    Else
      Call WBS_Option.clearCalendar
    End If
  Else
    Call WBS_Option.clearCalendar
  End If
  Call ProgressBar.showStart
  
  
  Call Import.ファイルインポート
  Call Calendar.書式設定
  Call Import.カレンダー用日程取得
  
  If setVal("lineColorFlg") = "True" Then
    setVal("lineColorFlg") = False
    Call WBS_Option.setLineColor
  Else
  End If
  
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
  Call WBS_Option.saveAndRefresh
  
  Application.Goto Reference:=Range("A6"), Scroll:=True


  Err.Clear
  Call Library.showNotice(200, "インポート")
End Sub


