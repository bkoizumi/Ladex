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
      
    LineColor = Library.getRegistry("Main", "LineColor")
    If LineColor = "" Then
      .LineColor.BackColor = 0
    Else
      .LineColor.BackColor = LineColor
    End If
  
    'Highlight設定
    HighLight_Color = Library.getRegistry("Main", "HighLight_Color")
    If HighLight_Color = "0" Then
      .HighLight_Color.BackColor = 10222585
    Else
      .HighLight_Color.BackColor = HighLight_Color
    End If
    .HighLight_Color.Caption = ""
    
    '透明度
    Highlight_TransparentRate = Library.getRegistry("Main", "Highlight_TransparentRate")
    If Highlight_TransparentRate = "0" Then
      .Highlight_TransparentRate.Value = 50
    Else
      .Highlight_TransparentRate.Value = Highlight_TransparentRate
    End If
  
    '表示方向
    Highlight_DspDirection = Library.getRegistry("Main", "Highlight_DspDirection")
    If Highlight_DspDirection = "X" Then
      Highlight_DspDirection_X.Value = True
      
    ElseIf Highlight_DspDirection = "Y" Then
      Highlight_DspDirection_Y.Value = True
    
    ElseIf Highlight_DspDirection = "B" Then
      Highlight_DspDirection_B.Value = True
    
    End If
  
    '表示方法
    Highlight_DspMethod = Library.getRegistry("Main", "Highlight_DspMethod")
    If Highlight_DspMethod = "0" Then
      Highlight_DspMethod_0.Value = True
    
    ElseIf Highlight_DspMethod = "0" Then
      Highlight_DspMethod_0.Value = True
      
    ElseIf Highlight_DspMethod = "1" Then
      Highlight_DspMethod_1.Value = True
    
    ElseIf Highlight_DspMethod = "2" Then
      Highlight_DspMethod_2.Value = True
    
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
Private Sub HighLight_Color_Click()
  Dim colorValue As Long
  
  colorValue = Library.getColor(Me.HighLight_Color.BackColor)
  Me.HighLight_Color.BackColor = colorValue
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
  
  
  Call Library.setRegistry("Main", "zoomLevel", Me.zoomLevel.Text)
  Call Library.setRegistry("Main", "gridLine", Me.gridLine.Value)
  Call Library.setRegistry("Main", "bgColor", Me.bgColor.Value)
  Call Library.setRegistry("Main", "LineColor", Me.LineColor.BackColor)
  
  
  Call Library.setRegistry("Main", "HighLight_Color", Me.HighLight_Color.BackColor)
  Call Library.setRegistry("Main", "Highlight_TransparentRate", Me.Highlight_TransparentRate.Value)

  '表示方向
  If Highlight_DspDirection_X.Value = True Then
    Highlight_DspDirection = "X"
  ElseIf Highlight_DspDirection_Y.Value = True Then
    Highlight_DspDirection = "Y"
  ElseIf Highlight_DspDirection_B.Value = True Then
    Highlight_DspDirection = "B"
  End If
  Call Library.setRegistry("Main", "Highlight_DspDirection", Highlight_DspDirection)



  '表示方向
  If Highlight_DspMethod_0.Value = True Then
    Highlight_DspMethod = "0"
  
  ElseIf Highlight_DspMethod_1.Value = True Then
    Highlight_DspMethod = "1"
  
  ElseIf Highlight_DspMethod_2.Value = True Then
    Highlight_DspMethod = "2"
  End If
  Call Library.setRegistry("Main", "Highlight_DspMethod", Highlight_DspMethod)




  
  
  'スタイルシートをスタイル2シートへコピー
'  endLine = sheetStyle2.Cells(Rows.count, 2).End(xlUp).Row
'  sheetStyle2.Range("A1:J" & endLine).Copy Destination:=sheetStyle.Range("A1")


  Unload Me
End Sub

  
