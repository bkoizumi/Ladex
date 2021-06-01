VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "オプション"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "Frm_Option.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret As Boolean




'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim zoomLevelVal  As Variant
  Dim setZoomLevel As String
  Dim endLine As Long
  Dim indexCnt As Integer
  
  Call init.setting
  Application.Cursor = xlDefault
  indexCnt = 0
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  
  With Frm_Option
    For Each zoomLevelVal In Split("25,50,75,85,100", ",")
      .zoomLevel.AddItem zoomLevelVal
      
      If zoomLevelVal = setZoomLevel Then
        .zoomLevel.ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .gridLine.Value = Library.getRegistry("Main", "gridLine")
    .bgColor.Value = Library.getRegistry("Main", "bgColor")
      
    If Library.getRegistry("Main", "highLightColor") = "" Then
      .highLightColor.BackColor = 10222585
    Else
      .highLightColor.BackColor = Library.getRegistry("Main", "highLightColor")
    End If
    
    If Library.getRegistry("Main", "LineColor") = "" Then
      .LineColor.BackColor = 0
    Else
      .LineColor.BackColor = Library.getRegistry("Main", "LineColor")
    End If
  End With
End Sub

'**************************************************************************************************
' * スタイル設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub IncludeFont01_Click()
  If IncludeFont01.Value = True Then
    ret = セルの書式設定_フォント(1)
     IncludeFont01.Value = ret
  End If
End Sub

'**************************************************************************************************
' * 組み込みダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function セルの書式設定_フォント(Optional line As Long = 1)
  Call init.setting
  sheetStyle2.Select
  sheetStyle2.Cells(line + 1, 11).Select
  ret = Application.Dialogs(xlDialogActiveCellFont).Show
  If ret = True Then
    sheetStyle2.Cells(line + 1, 5) = "TRUE"
  Else
    sheetStyle2.Cells(line + 1, 5) = "FALSE"
  End If
  セルの書式設定_フォント = ret
End Function















'**************************************************************************************************
' * 基本設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
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
' * 処理キャンセル
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()

  Call Library.setRegistry("UserForm", "OptionTop", Me.Top)
  Call Library.setRegistry("UserForm", "OptionLeft", Me.Left)
  
  Unload Me
End Sub


'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  Dim execDay As Date
  
  Call Library.setRegistry("UserForm", "OptionTop", Me.Top)
  Call Library.setRegistry("UserForm", "OptionLeft", Me.Left)
  
  
  Call Library.setRegistry("Main", "zoomLevel", Me.zoomLevel.Text)
  Call Library.setRegistry("Main", "gridLine", Me.gridLine.Value)
  Call Library.setRegistry("Main", "bgColor", Me.bgColor.Value)
  Call Library.setRegistry("Main", "highLightColor", Me.highLightColor.BackColor)
  Call Library.setRegistry("Main", "LineColor", Me.LineColor.BackColor)
  
  'スタイルシートをスタイル2シートへコピー
'  endLine = sheetStyle2.Cells(Rows.count, 2).End(xlUp).Row
'  sheetStyle2.Range("A1:J" & endLine).Copy Destination:=sheetStyle.Range("A1")


  Unload Me
End Sub

