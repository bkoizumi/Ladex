VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'https://excel-ubara.com/vba_class/VBA_CLASS_02.html

Option Explicit

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
    If Me.StopEvent Then Exit Sub
    If Not CheckBook(wb) Then Exit Sub

    'メッセージボックスでイベント補足を確認
'    MsgBox "WorkbookOpen : " & wb.FullName
End Sub


'==================================================================================================
'ブックが閉じられる直前に発生
Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    If Me.StopEvent Then Exit Sub
    If Not CheckBook(wb) Then Exit Sub

    'メッセージボックスでイベント補足を確認
'    MsgBox "WorkbookBeforeClose : " & wb.FullName
End Sub


'==================================================================================================
'新しいブックが作成されたときに発生
Private Sub xlApp_NewWorkbook(ByVal wb As Workbook)
    If Me.StopEvent Then Exit Sub
    If Not CheckBook(wb) Then Exit Sub

    'メッセージボックスでイベント補足を確認
'    MsgBox "NewWorkbook : " & wb.Name
End Sub

'==================================================================================================
' Applicationのシートイベント
'==================================================================================================


'==================================================================================================
'シートがアクティブになったときに発生
Private Sub xlApp_SheetActivate(ByVal sh As Object)
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub

    'メッセージボックスでイベント補足を確認
'    MsgBox "SheetActivate : [" & Sh.Parent.Name & "]" & Sh.Name
  Exit Sub
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, "", True)
End Sub


'==================================================================================================
'ワークシートのセルが変更されたときに発生
Private Sub xlApp_SheetChange(ByVal sh As Object, ByVal Target As Range)
    If Me.StopEvent Then Exit Sub
    If Not CheckSheet(sh) Then Exit Sub



    'メッセージボックスでイベント補足を確認
'    MsgBox "SheetChange : " & Target.Address(external:=True)
End Sub


'==================================================================================================
Private Sub xlApp_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub
  
  'ハイライト機能----------------------------------------------------------------------------------
  If Library.getRegistry("Main", "HighLightFlg") = True Then
    runFlg = True
    Call Ctl_HighLight.showStart(Target)
  End If
  
  '数式確認機能------------------------------------------------------------------------------------
  If BKcf_rbPressed = True Then
    runFlg = True
    Call Ctl_Formula.数式確認
  End If
End Sub


'==================================================================================================
'セルがダブルクリックされたときに発生
Private Sub xlApp_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub

  'ズーム機能--------------------------------------------------------------------------------------
  If Library.getRegistry("Main", "ZoomFlg") = True Then
    Cancel = True
    runFlg = True
    Call Ctl_Zoom.ZoomIn(CStr(Target.Address(False, False)))
  End If




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
