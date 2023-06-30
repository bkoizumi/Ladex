VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Style 
   Caption         =   "スタイル管理"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675.001
   OleObjectBlob   =   "Frm_Style.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public InitializeFlg  As Boolean
Public selectLine     As Long







'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim dicKey As Variant
  
  Const funcName As String = "Frm_Style.UserForm_Initialize"

  '処理開始--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '表示位置指定----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
    
  InitializeFlg = True
  
  For Each dicKey In useStyleVal
    styleList.AddItem dicKey
  Next
  styleList.ListIndex = 0

  
    
  InitializeFlg = False
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub



'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub Btn_Delete_Click()
  Dim dicKey As Variant
  
  ActiveWorkbook.Styles(styleList.Value).delete
  useStyleVal.Remove (styleList.Value)
  styleList.Clear
  
  For Each dicKey In useStyleVal
    If dicKey <> "" Then
      styleList.AddItem dicKey
    End If
  Next
  On Error Resume Next
  styleList.ListIndex = 0
  
End Sub


'==================================================================================================
Private Sub styleList_Change()
  Dim objSheetName
  Dim slctRange As String
  Dim styleName As String, sheetName As String
  Dim objShp
  Const funcName As String = "Frm_Style.styleList_Change"

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  styleName = styleList.Value
  
  SheetList.ListItems.Clear
  SheetList.ColumnHeaders.Clear


  '既存のオブジェクト削除
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmStyleName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  
  
  With SheetList
    .View = lvwReport
    .LabelEdit = lvwManual
    .HideSelection = False
    .AllowColumnReorder = True
    .FullRowSelect = True
    .Gridlines = True
    .ColumnHeaders.add , "_SheetName", "シート名", 140
    .ColumnHeaders.add , "_Address", "セル", 300
    
    
    Call Library.showDebugForm("styleName", useStyleVal(styleName), "debug")
    For Each objSheetName In Split(useStyleVal(styleName), "<|>")
      sheetName = Split(objSheetName, "!")(0)
      slctRange = Split(objSheetName, "!")(1)
      
      With .ListItems.add
        .Text = sheetName
        .SubItems(1) = slctRange
      End With
    Next
  End With
End Sub



'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Dim objShp
  Set useStyleVal = Nothing
  
  '既存のオブジェクト削除
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmStyleName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  Unload Me
End Sub

'==================================================================================================
Private Sub SheetList_Click()
  Dim sheetName As String, slctRange As String
  Dim objShp, objSlctRange
  
  
  Const funcName As String = "Frm_Style.SheetList_Click"

  '処理開始--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  '既存のオブジェクト削除
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "confirmStyleName_*" Then
      ActiveSheet.Shapes(objShp.Name).delete
    End If
  Next
  
  
  sheetName = SheetList.SelectedItem
  If sheetName = "Abort" Then
    Exit Sub
  End If
  
  slctRange = SheetList.SelectedItem.SubItems(1)
  
  
  
  ActiveWorkbook.Worksheets(sheetName).Select
  
  
  '選択範囲に枠をつける
  For Each objSlctRange In Split(slctRange, ",")
    With ActiveSheet.Range(objSlctRange)
      ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height).Select
    End With
    Selection.Name = "confirmStyleName_" & styleList.Value & objSlctRange
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(205, 205, 255)
    Selection.ShapeRange.Fill.Transparency = 0.5
    Selection.Text = styleList.Value
    
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = "メイリオ"
      .TextRange.Font.NameFarEast = "メイリオ"
      .TextRange.Font.Name = "メイリオ"
      .TextRange.Font.Size = 9
      .MarginLeft = 3
      .MarginRight = 0
      .MarginTop = 0
      .MarginBottom = 0
      .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    Selection.ShapeRange.line.Visible = msoTrue
    Selection.ShapeRange.line.ForeColor.RGB = RGB(255, 0, 0)
    Selection.ShapeRange.line.Weight = 2
  Next
  Range(slctRange).Select
  
  Exit Sub

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

