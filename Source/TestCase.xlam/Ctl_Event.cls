VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ctl_Event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'https://excel-ubara.com/vba_class/VBA_CLASS_02.html

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
  Dim wbsPath As String
  Const funcName As String = "Ctl_Event.xlApp_WorkbookOpen"
  
  Dim scrollVal As Long
  
  'Stop
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  Call init.setting
'  If wb.Name Like "*WBS*.xlsm" And Library.chkBookOpened("WBS.xlam") = False Then
'    If Library.chkFileExists(LadexDir & "\WBS.xlam") = True Then
'      wbsPath = LadexDir & "\WBS.xlam"
'    ElseIf Library.chkFileExists("C:\WorkSpace\000.myProject\ExcelMacro\WBS\WBS.xlam") = True Then
'      wbsPath = "C:\WorkSpace\000.myProject\ExcelMacro\WBS\WBS.xlam"
'    End If
'
'    runFlg = True
'    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
'    'Workbooks.Open "C:\WorkSpace\myProject\ExcelMacro\WBS\WBS.xlam"
'    Workbooks.Open wbsPath
'    Call Library.endScript
'  End If

  
  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub


'==================================================================================================
'ブックが閉じられる直前に発生
Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_WorkbookBeforeClose"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Ctl_System.resetScroll
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
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
  Const funcName As String = "Ctl_Event.xlApp_SheetActivate"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(Sh) Then Exit Sub

'  runFlg = True
'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'
'  Call Ctl_Sheet.スクロール設定
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
'ワークシートのセルが変更されたときに発生
Private Sub xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  Const funcName As String = "Ctl_Event.xlApp_SheetChange"
    
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(Sh) Then Exit Sub

'  Call init.setting
'  If Sh.Name = "" And Target.count = 1 Then
'    Call Ctl_TestCase.書式設定(Target, True)
'  End If


  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub


'==================================================================================================
'ワークシートが変更されたとき
Private Sub xlApp_SheetDeactivate(ByVal Sh As Object)
  Const funcName As String = "Ctl_Event.xlApp_SheetDeactivate"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(Sh) Then Exit Sub

  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
'ワークシートの選択セルが変更されたときに発生
Private Sub xlApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  Dim slctCells As Range
  
  Const funcName As String = "Ctl_Event.xlApp_SheetSelectionChange"
 
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(Sh) Then Exit Sub
  

  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
   Call Library.endScript
End Sub


'==================================================================================================
'セルがダブルクリックされたときに発生
Private Sub xlApp_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_SheetBeforeDoubleClick"

  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(Sh) Then Exit Sub

 
  Exit Sub
  
'エラー発生時------------------------------------
catchError:
  'Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
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
