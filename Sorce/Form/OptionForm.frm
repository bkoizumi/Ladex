VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionForm 
   Caption         =   "オプション"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "OptionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "OptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub highLightColor_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.highLightColor.BackColor)
  Me.highLightColor.BackColor = colorValue
End Sub


Private Sub LineColor_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.LineColor.BackColor)
  Me.LineColor.BackColor = colorValue
End Sub

'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim zoomLevelVal  As Variant
  Dim setZoomLevel As String
  
  Dim indexCnt As Integer
  
  Application.Cursor = xlDefault
  indexCnt = 0
  setZoomLevel = Library.getRegistry("zoomLevel")
  
  With OptionForm
    For Each zoomLevelVal In Split("25,50,75,85,100", ",")
      .zoomLevel.AddItem zoomLevelVal
      
      If zoomLevelVal = setZoomLevel Then
        .zoomLevel.ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .gridLine.Value = Library.getRegistry("gridLine")
    .bgColor.Value = Library.getRegistry("bgColor")
    .highLightColor.BackColor = Library.getRegistry("highLightColor")
    .LineColor.BackColor = Library.getRegistry("LineColor")
  End With
End Sub

'**************************************************************************************************
' * 処理キャンセル
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()
  Unload Me
End Sub


'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  Dim execDay As Date
  
  Call Library.setRegistry("zoomLevel", Me.zoomLevel.Text)
  Call Library.setRegistry("gridLine", Me.gridLine.Value)
  Call Library.setRegistry("bgColor", Me.bgColor.Value)
  Call Library.setRegistry("highLightColor", Me.highLightColor.BackColor)
  Call Library.setRegistry("LineColor", Me.LineColor.BackColor)
  
  Unload Me
End Sub

