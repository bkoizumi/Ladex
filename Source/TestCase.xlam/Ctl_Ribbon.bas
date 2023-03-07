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
'  ribbonUI.ActivateTab "TestCaseTab"

End Function


'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '処理開始--------------------------------------
'  runFlg = True
'  On Error GoTo catchError
'  Call init.setting(True)
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.startScript
'  PrgP_Max = 4
'  resetCellFlg = True
'  runFlg = True
'  Call Ctl_ProgressBar.showStart
  '----------------------------------------------

  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "CloseAddOn"
      Call Library.addonClose
    
    Case "OptionAddin解除"
      Workbooks(ThisWorkbook.Name).IsAddin = False
    
    'テスト仕様書--------------------------------
    Case "シート追加"
      Call Ctl_TestCase.シート追加
      
    Case "書式設定"
      Call Ctl_TestCase.再設定
      
      
    'selenium------------------------------------
    Case "Selenium実行"
      Call Ctl_Selenium.開始
      
      
      
      
      
    Case Else
      Call Library.showDebugForm("リボンメニューなし", control.ID, "Error")
      Call Library.showNotice(406, "リボンメニューなし：" & control.ID, True)
  End Select
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
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
