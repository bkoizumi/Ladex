VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ctl_Event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 And Win64 Then
  Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As LongPtr) As Long

#Else
  Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long
#End If


'対象のExcelアプリケーション
Private WithEvents xlApp As Excel.Application
Attribute xlApp.VB_VarHelpID = -1

'対象ブック・シート：1次元配列の要素に構造体を入れる
Private Type typeBookSheets
    Book As String 'ブック
    Sheets() As String 'シート1,シート2,…
End Type
Private pBookSheets() As typeBookSheets
Private pCntBS As Integer

'イベント停止（一時的に停止させる）
Private pStopEvent As Boolean


'==================================================================================================
' 公開プロパティ
'==================================================================================================
'Excelアプリケーションの設定・取得
Public Property Set ExcelApplication(argApp As Excel.Application)
    Set xlApp = argApp
End Property

Public Property Get ExcelApplication() As Excel.Application
    Set ExcelApplication = xlApp
End Property

'イベント停止（一時的に停止させる）の取得設定
Public Property Let StopEvent(argStop As Boolean)
    pStopEvent = argStop
End Property

Public Property Get StopEvent() As Boolean
  StopEvent = pStopEvent
End Property

'==================================================================================================
' 公開メソッド
'==================================================================================================
'対象ブック・シートの初期化
Public Sub InitializeBookSheets()
    pCntBS = 0
    Erase pBookSheets
End Sub


'==================================================================================================
'対象ブック・シートの設定
Public Function AddBookSheet(ByRef argAry) As Boolean
    Dim i1 As Long
    Dim i2 As Long
    Dim ix As Integer
    Dim myArray() As String

    '配列の次元数により分岐
    Select Case GetDimension(argAry)

        Case 1 '1次元配列
            ReDim Preserve pBookSheets(pCntBS)
            pBookSheets(pCntBS).Book = argAry(LBound(argAry))
            ix = 0
            ReDim myArray(ix)
            For i1 = LBound(argAry) + 1 To UBound(argAry)
                If argAry(i1) <> "" Then
                    ReDim Preserve myArray(ix)
                    myArray(ix) = argAry(i1)
                    ix = ix + 1
                End If
            Next
            pBookSheets(pCntBS).Sheets = myArray
            Erase myArray
            pCntBS = pCntBS + 1

        Case 2 '2次元配列
            For i1 = LBound(argAry, 1) To UBound(argAry, 1)
                ReDim Preserve pBookSheets(pCntBS)
                pBookSheets(pCntBS).Book = argAry(i1, LBound(argAry, 2))
                ix = 0
                ReDim myArray(ix)
                For i2 = LBound(argAry, 2) + 1 To UBound(argAry, 2)
                    If argAry(i1, i2) <> "" Then
                        ReDim Preserve myArray(ix)
                        myArray(ix) = argAry(i1, i2)
                        ix = ix + 1
                    End If
                Next
                pBookSheets(pCntBS).Sheets = myArray
                Erase myArray
                pCntBS = pCntBS + 1
            Next

        Case Else '1,2次元以外はエラーで戻す
            AddBookSheet = False
            Exit Function
    End Select

    AddBookSheet = True
End Function


'==================================================================================================
'対象ブック・シートの全取得：2次元配列で戻す
Public Function GetBookSheets() As Variant
    Dim i1 As Integer, i2 As Integer
    Dim intMax As Integer
    Dim myArray() As String

    '配列内の最大シート数を取得
    intMax = 0
    For i1 = 0 To UBound(pBookSheets)
        If UBound(pBookSheets(i1).Sheets) > intMax Then
            intMax = UBound(pBookSheets(i1).Sheets)
        End If
    Next

    '構造体の1次元配列を、構造体をバラして2次元配列に
    ReDim myArray(UBound(pBookSheets), intMax + 1)
    For i1 = 0 To UBound(pBookSheets)
        myArray(i1, 0) = pBookSheets(i1).Book
        For i2 = LBound(pBookSheets(i1).Sheets) To UBound(pBookSheets(i1).Sheets)
            myArray(i1, i2 + 1) = pBookSheets(i1).Sheets(i2)
        Next
    Next
    GetBookSheets = myArray
End Function

'==================================================================================================
' クラスのメソッド
'==================================================================================================

'クラスの初期化
Private Sub Class_Initialize()
    Set Me.ExcelApplication = Application
    Me.StopEvent = False
    Me.InitializeBookSheets
End Sub

'クラスの終了処理
Private Sub Class_Terminate()
    Set xlApp = Nothing
End Sub



'==================================================================================================
' Applicationのブックイベント
'==================================================================================================


'==================================================================================================
'ブックが開かれたときに発生
Private Sub xlApp_WorkbookOpen(ByVal wb As Workbook)
  Dim macroPath As String
  Const funcName As String = "Ctl_Event.xlApp_WorkbookOpen"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  '----------------------------------------------

  macroPath = ""
  Call init.setting
  If wb.Name Like "*WBS*.xlsm" And Library.chkBookOpened("WBS.xlam") = False Then
    If Library.chkFileExists(AddInDir & "\WBS.xlam") = True Then
      macroPath = AddInDir & "\WBS.xlam"
    End If
  
  ElseIf wb.Name Like "*テスト仕様書兼結果報告書*.xlsx" And Library.chkBookOpened("TestCase.xlam") = False Then
    If Library.chkFileExists(AddInDir & "\TestCase.xlam") = True Then
      macroPath = AddInDir & "\TestCase.xlam"
    End If
  End If
  
  If macroPath <> "" Then
    Call Library.showDebugForm("macroPath", macroPath, "debug")
    Workbooks.Open macroPath
  End If

  
  '処理終了--------------------------------------
  Call Library.endScript
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end1")
  '----------------------------------------------
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
'ブックが閉じられる直前に発生
Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
  Dim macroPath  As String
  Dim chkOpenFile As Variant, chkOpenFileCnt As Integer
  
  Const funcName As String = "Ctl_Event.xlApp_WorkbookBeforeClose"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  '----------------------------------------------

  If wb.Name Like "*WBS*.xlsm" And Library.chkBookOpened("WBS.xlam") = True Then
    macroPath = "WBS.xlam"
    chkOpenFile = Library.chkIsOpen("*WBS*", chkOpenFileCnt)
  
  ElseIf wb.Name Like "*テスト仕様書兼結果報告書*.xlsx" And Library.chkBookOpened("TestCase.xlam") = True Then
    macroPath = "TestCase.xlam"
    chkOpenFile = Library.chkIsOpen("*テスト仕様書兼結果報告書*", chkOpenFileCnt)

  End If
  
  If macroPath <> "" And chkOpenFileCnt = 1 Then
    Call Library.showDebugForm("macroPath", macroPath, "debug")
    
    Select Case MsgBox("'" & wb.Name & "' の変更内容を保存しますか？", vbYesNoCancel + vbQuestion + vbDefaultButton3, "")
      Case vbYes
        ActiveWorkbook.Close SaveChanges:=True
        Workbooks(macroPath).Close
      Case vbNo
        ActiveWorkbook.Close SaveChanges:=False
        Workbooks(macroPath).Close
      Case vbCancel
        Cancel = True
      Case Else
    End Select
  End If
  
  
  '処理終了--------------------------------------
  Call Library.endScript
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end1")
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'==================================================================================================
'ブックが非アクティブになったとき
'Private Sub xlApp_SheetDeactivate()
'  Const funcName As String = "Ctl_Event.xlApp_SheetDeactivate"
'
'  On Error GoTo catchError
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
'
'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'
'  Call Ctl_System.resetScroll
'
'  Call Library.showDebugForm(funcName, , "end")
'
'  Exit Sub
'
''エラー発生時------------------------------------
'catchError:
'  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
'End Sub


'==================================================================================================
'新しいブックが作成されたときに発生
Private Sub xlApp_NewWorkbook(ByVal wb As Workbook)
  Const funcName As String = "Ctl_Event.xlApp_NewWorkbook"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
' Applicationのシートイベント
'==================================================================================================


'==================================================================================================
'シートがアクティブになったときに発生
Private Sub xlApp_SheetActivate(ByVal Sh As Object)
  Dim wb As Workbook
  Dim macroPath As String
  
  Const funcName As String = "Ctl_Event.xlApp_SheetActivate"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  If Application.CutCopyMode Then Exit Sub
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  '----------------------------------------------
  
  Set wb = ActiveWorkbook
  macroPath = ""
  Call init.setting
  If wb.Name Like "*WBS*.xlsm" And Library.chkBookOpened("WBS.xlam") = False Then
    If Library.chkFileExists(AddInDir & "\WBS.xlam") = True Then
      macroPath = AddInDir & "\WBS.xlam"
    End If
  
  ElseIf wb.Name Like "*テスト仕様書兼結果報告書*.xlsx" And Library.chkBookOpened("TestCase.xlam") = False Then
    If Library.chkFileExists(AddInDir & "\TestCase.xlam") = True Then
      macroPath = AddInDir & "\TestCase.xlam"
    End If
    
  End If
  
  If macroPath <> "" Then
    runFlg = True
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.showDebugForm("macroPath", macroPath, "debug")
    Workbooks.Open macroPath
  End If
  
  '処理終了--------------------------------------
  Call Library.endScript
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end1")
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
'ワークシートのセルが変更されたときに発生
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  Const funcName As String = "Ctl_Event.xlApp_SheetChange"
    
  '処理開始--------------------------------------
  On Error GoTo catchError
  If Application.CutCopyMode Then Exit Sub
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  
'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start1")
'  Call Library.startScript
  '----------------------------------------------


  '処理終了--------------------------------------
'  Call Library.endScript
'  Call init.resetGlobalVal
'  Call Library.showDebugForm(funcName, , "end1")
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
'ワークシートが変更されたとき
Private Sub xlApp_SheetDeactivate(ByVal Sh As Object)
  Const funcName As String = "Ctl_Event.xlApp_SheetDeactivate"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  If Application.CutCopyMode Then Exit Sub
  '----------------------------------------------
  
  'ハイライト機能--------------------------------
  If BKh_rbPressed = True Then
    Call init.setting
    Call Library.startScript
    If Library.chkShapeName("HighLight_X", Sh) = True Then
      Sh.Shapes("HighLight_X").delete
    End If
    If Library.chkShapeName("HighLight_Y", Sh) = True Then
      Sh.Shapes("HighLight_Y").delete
    End If
      
    Call Library.setRegistry("targetInfo", "HighLight_Book", ActiveWorkbook.Name)
    Call Library.setRegistry("targetInfo", "HighLight_Sheet", ActiveSheet.Name)
    
    runFlg = True
    Call Ctl_HighLight.showStart(ActiveCell, True)
    Call Library.endScript
  End If
  
  
  '処理終了--------------------------------------
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub

'==================================================================================================
'ワークシートの選択セルが変更されたときに発生
Private Sub xlApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Event.xlApp_SheetSelectionChange"
  
  '処理開始--------------------------------------
  On Error GoTo catchError
  If Application.CutCopyMode Then Exit Sub
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  '----------------------------------------------
  
  'ハイライト機能--------------------------------
  If BKh_rbPressed = True Then
    If Library.getRegistry("targetInfo", "HighLight_Book", "String") = ActiveWorkbook.Name Then
      runFlg = True
      
      Call Library.showDebugForm("Target ", Target.Address, "debug")
      Call Ctl_HighLight.showStart(Target, True)
    
    Else
      Call init.setting
      Call Library.startScript
      If Library.chkShapeName("HighLight_X", Sh) = True Then
        Sh.Shapes("HighLight_X").delete
      End If
      If Library.chkShapeName("HighLight_Y", Sh) = True Then
        Sh.Shapes("HighLight_Y").delete
      End If
      Call Library.endScript
    End If
  
  End If
  
  '数式確認機能----------------------------------
  If BKcf_rbPressed = True Then
    If Library.getRegistry("targetInfo", "Formula_Book", "String") = ActiveWorkbook.Name And Library.getRegistry("targetInfo", "Formula_Sheet", "String") = ActiveSheet.Name Then
      runFlg = True
      Call Ctl_Formula.数式確認
    End If
  End If
  
  '処理終了--------------------------------------
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'==================================================================================================
'セルがダブルクリックされたときに発生
Private Sub xlApp_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_SheetBeforeDoubleClick"

  '処理開始--------------------------------------
  On Error GoTo catchError
'  If Me.StopEvent Then Exit Sub
'  If Not CheckBook(wb) Then Exit Sub
  '----------------------------------------------

  
  'ズーム機能------------------------------------
  If BKz_rbPressed = True Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    
    Cancel = True
    runFlg = True
    Call Ctl_Zoom.ZoomIn(CStr(Target.Address(False, False)))
    
    Call Library.showDebugForm(funcName, , "end")
  End If
  
  '処理終了--------------------------------------
  Exit Sub
  '----------------------------------------------

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'==================================================================================================
'右クリックされたとき
Private Sub xlApp_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_SheetBeforeRightClick"

'  On Error GoTo catchError
'  If Me.StopEvent Then Exit Sub
'  If Not CheckSheet(Sh) Then Exit Sub
'
'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'
'  '右クリックメニュー----------------------------
'  If Library.getRegistry("Main", "CustomRibbon", "Boolean") = True Then
'    Call Main.右クリックメニュー(Target, Cancel)
'  End If
'
'  Call Library.showDebugForm(funcName, , "end")
'  Exit Sub

'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub


'==================================================================================================
' 非公開メソッド
'==================================================================================================


'==================================================================================================
'対象ブックの判定
Private Function CheckBook(ByVal wb As Workbook) As Boolean
  '指定なしはすべて対象とする
  If pCntBS = 0 Then
    CheckBook = True
    Exit Function
  End If
  
  'ブック・シート指定の配列にブック名があるか判定
  Dim i1 As Integer
  CheckBook = False
  For i1 = 0 To UBound(pBookSheets)
    'ブックの名称判定
    If wb.Name Like pBookSheets(i1).Book Then
      'シート指定がない時のみブックの指定と判定
      If pBookSheets(i1).Sheets(0) = "" Then
        CheckBook = True
        Exit Function
      End If
    End If
  Next
End Function

'==================================================================================================
'対象シートの判定
Private Function CheckSheet(ByVal ws As Worksheet) As Boolean
    
  '指定なしはすべて対象とする
  If pCntBS = 0 Then
    CheckSheet = True
    Exit Function
  End If

  'ブック・シート指定の配列にシート名があるか判定
  Dim i1 As Integer, i2 As Integer
  Dim wb As Workbook
  CheckSheet = False
  Set wb = ws.Parent
  For i1 = 0 To UBound(pBookSheets)
    'ブックの名称判定
    If wb.Name Like pBookSheets(i1).Book Then
      'シートの名称判定
      For i2 = LBound(pBookSheets(i1).Sheets) To UBound(pBookSheets(i1).Sheets)
        If ws.Name Like pBookSheets(i1).Sheets(i2) Then
          CheckSheet = True
          Exit Function
        End If
      Next
    End If
  Next
End Function


'==================================================================================================
'配列の次元数取得とエラー値判定：1,2以外の戻り値はエラー
Private Function GetDimension(ByVal argAry) As Integer
  On Error Resume Next
  
  '配列の次元数取得
  Dim intDim As Integer
  Dim tmpDim As Integer
  intDim = 0
  Do While Err.Number = 0
      intDim = intDim + 1
      'エラーになった時点で配列が終わり
      tmpDim = UBound(argAry, intDim)
  Loop
  Err.Clear
  intDim = intDim - 1
  
  '配列の場合、エラー値データの判定
  Dim vTemp
  Select Case intDim
    Case 1, 2
      For Each vTemp In argAry
        If IsError(vTemp) Then
          GetDimension = 0
          Exit Function
        End If
      Next
  End Select
  
  GetDimension = intDim
  On Error GoTo 0
End Function
