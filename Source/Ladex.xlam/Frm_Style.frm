VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Style 
   Caption         =   "�X�^�C���Ǘ�"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675.001
   OleObjectBlob   =   "Frm_Style.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim dicKey As Variant
  
  Const funcName As String = "Frm_Style.UserForm_Initialize"

  '�����J�n--------------------------------------
'  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------
  
  '�\���ʒu�w��----------------------------------
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

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub



'**************************************************************************************************
' * �{�^������������
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

  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  styleName = styleList.Value
  
  SheetList.ListItems.Clear
  SheetList.ColumnHeaders.Clear


  '�����̃I�u�W�F�N�g�폜
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
    .ColumnHeaders.add , "_SheetName", "�V�[�g��", 140
    .ColumnHeaders.add , "_Address", "�Z��", 300
    
    
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
'�L�����Z������
Private Sub Cancel_Click()
  Dim objShp
  Set useStyleVal = Nothing
  
  '�����̃I�u�W�F�N�g�폜
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

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  '�����̃I�u�W�F�N�g�폜
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
  
  
  '�I��͈͂ɘg������
  For Each objSlctRange In Split(slctRange, ",")
    With ActiveSheet.Range(objSlctRange)
      ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, Top:=.Top, Width:=.Width, Height:=.Height).Select
    End With
    Selection.Name = "confirmStyleName_" & styleList.Value & objSlctRange
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(205, 205, 255)
    Selection.ShapeRange.Fill.Transparency = 0.5
    Selection.Text = styleList.Value
    
    With Selection.ShapeRange.TextFrame2
      .TextRange.Font.NameComplexScript = "���C���I"
      .TextRange.Font.NameFarEast = "���C���I"
      .TextRange.Font.Name = "���C���I"
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

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
End Sub

