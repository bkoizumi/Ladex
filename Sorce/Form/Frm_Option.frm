VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "オプション"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
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
  
  setZoomLevel = Library.getRegistry("Main", "ZoomLevel")
  
  With Frm_Option
    For Each zoomLevelVal In Split("25,50,75,85,100", ",")
      .ZoomLevel.AddItem zoomLevelVal
      
      If zoomLevelVal = setZoomLevel Then
        .ZoomLevel.ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .GridLine.Value = Library.getRegistry("Main", "gridLine")
    .BgColor.Value = Library.getRegistry("Main", "bgColor")
      
    LineColor = Library.getRegistry("Main", "LineColor")
    If LineColor = "" Then
      .LineColor.BackColor = 0
    Else
      .LineColor.BackColor = LineColor
    End If
    .LineColor.Caption = ""
    
  
    'Highlight設定
    HighLightColor = Library.getRegistry("Main", "HighLightColor")
    If HighLightColor = "0" Then
      .HighLightColor.BackColor = 10222585
    Else
      .HighLightColor.BackColor = HighLightColor
    End If
    .HighLightColor.Caption = ""
    
    '透明度
    HighlightTransparentRate = Library.getRegistry("Main", "HighLightTransparentRate")
    If HighlightTransparentRate = "0" Then
      .HighlightTransparentRate.Value = 50
    Else
      .HighlightTransparentRate.Value = HighlightTransparentRate
    End If
  
    '表示方向
    HighLightDspDirection = Library.getRegistry("Main", "HighLightDspDirection")
    If HighLightDspDirection = "X" Then
      HighlightDspDirection_X.Value = True
      
    ElseIf HighLightDspDirection = "Y" Then
      Highlight_DspDirection_Y.Value = True
    
    ElseIf Highlight_DspDirection = "B" Then
      Highlight_DspDirection_B.Value = True
    
    End If
  
    '表示方法
    HighLightDspMethod = Library.getRegistry("Main", "HighLightDspMethod")
    If HighLightDspMethod = "0" Then
      HighlightDspMethod_0.Value = True
    
    ElseIf HighLight_DspMethod = "0" Then
      HighlightDspMethod_0.Value = True
      
    ElseIf HighLight_DspMethod = "1" Then
      HighlightDspMethod_1.Value = True
    
    ElseIf HighLight_DspMethod = "2" Then
      HighlightDspMethod_2.Value = True
    
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
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub HighLightColor_Click()
  Dim colorValue As Long
  
  colorValue = Library.getColor(Me.HighLightColor.BackColor)
  Me.HighLightColor.BackColor = colorValue
End Sub


'==================================================================================================
Private Sub LineColor_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.LineColor.BackColor)
  Me.LineColor.BackColor = colorValue
End Sub


'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()

  Call Library.setRegistry("UserForm", "OptionTop", Me.Top)
  Call Library.setRegistry("UserForm", "OptionLeft", Me.Left)
  
  Unload Me
End Sub


'==================================================================================================
' 実行
Private Sub run_Click()
  Dim execDay As Date
  
  Call Library.setRegistry("UserForm", "OptionTop", Me.Top)
  Call Library.setRegistry("UserForm", "OptionLeft", Me.Left)
  
  
  Call Library.setRegistry("Main", "ZoomLevel", Me.ZoomLevel.Text)
  Call Library.setRegistry("Main", "GridLine", Me.GridLine.Value)
  Call Library.setRegistry("Main", "bgColor", Me.BgColor.Value)
  Call Library.setRegistry("Main", "LineColor", Me.LineColor.BackColor)
  
  
  Call Library.setRegistry("Main", "HighLightColor", Me.HighLightColor.BackColor)
  Call Library.setRegistry("Main", "HighLightTransparentRate", Me.HighlightTransparentRate.Value)

  '表示方向
  If HighlightDspDirection_X.Value = True Then
    HighLightDspDirection = "X"
  ElseIf Highlight_DspDirection_Y.Value = True Then
    HighLightDspDirection = "Y"
  ElseIf HighlightDspDirection_B.Value = True Then
    HighLightDspDirection = "B"
  End If
  Call Library.setRegistry("Main", "HighLightDspDirection", HighLightDspDirection)



  '表示方向
  If HighlightDspMethod_0.Value = True Then
    HighLightDspMethod = "0"
  
  ElseIf HighlightDspMethod_1.Value = True Then
    HighLightDspMethod = "1"
  
  ElseIf HighlightDspMethod_2.Value = True Then
    HighLightDspMethod = "2"
  End If
  Call Library.setRegistry("Main", "HighLightDspMethod", HighLightDspMethod)




  
  
  'スタイルシートをスタイル2シートへコピー
'  endLine = sheetStyle2.Cells(Rows.count, 2).End(xlUp).Row
'  sheetStyle2.Range("A1:J" & endLine).Copy Destination:=sheetStyle.Range("A1")


  Unload Me
End Sub

  
