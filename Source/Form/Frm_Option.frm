VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "オプション"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "Frm_Option.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret As Boolean
Dim colorValue As Long
Dim HighLightDspDirection As String
Dim old_BKh_rbPressed  As Boolean
Public InitializeFlg   As Boolean



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
  Dim previewImgPath As String
  
  InitializeFlg = True
  
  Call init.setting
  Application.Cursor = xlDefault
  indexCnt = 0
  old_BKh_rbPressed = BKh_rbPressed
  
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
    
  
    'Highlight設定---------------------------------------------------------------------------------
    HighLightColor = Library.getRegistry("Main", "HighLightColor")
    If HighLightColor = "0" Then
      .HighLightColor.BackColor = 10222585
    Else
      .HighLightColor.BackColor = HighLightColor
    End If
    .HighLightColor.Caption = ""
    
    
    '透明度----------------------------------------------------------------------------------------
    .HighlightTransparentRate.Min = 0
    .HighlightTransparentRate.Max = 100
    HighlightTransparentRate = Library.getRegistry("Main", "HighLightTransparentRate")
    If HighlightTransparentRate = "0" Then
      .HighlightTransparentRate.Value = 70
      .HighlightTransparentRate_text.Caption = 70
    Else
      .HighlightTransparentRate.Value = HighlightTransparentRate
      .HighlightTransparentRate_text.Caption = HighlightTransparentRate
    End If
  
    '表示方向--------------------------------------------------------------------------------------
    HighLightDspDirection = Library.getRegistry("Main", "HighLightDspDirection")
    If HighLightDspDirection = "X" Then
      .HighlightDspDirection_X.Value = True
      
    ElseIf HighLightDspDirection = "Y" Then
      .HighlightDspDirection_Y.Value = True
    
    ElseIf HighLightDspDirection = "B" Then
      .HighlightDspDirection_B.Value = True
    
    End If
  
    '表示方法--------------------------------------------------------------------------------------
    HighLightDspMethod = Library.getRegistry("Main", "HighLightDspMethod")
    If HighLightDspMethod = "0" Then
      .HighlightDspMethod_0.Value = True
    
    ElseIf HighLightDspMethod = "0" Then
      .HighlightDspMethod_0.Value = True
      
    ElseIf HighLightDspMethod = "1" Then
      .HighlightDspMethod_1.Value = True
    
    ElseIf HighLightDspMethod = "2" Then
      .HighlightDspMethod_2.Value = True
    
    End If
    
    imageName = thisAppName & "HighLightImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoHighLightImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    End If
    HighLightImg.Picture = LoadPicture(previewImgPath)
    
    
    
    'コメント 背景色-------------------------------------------------------------------------------
    CommentBgColor = Library.getRegistry("Main", "CommentBgColor")
    If CommentBgColor = "0" Then
      .CommentColor.BackColor = 10222585
    Else
      .CommentColor.BackColor = CommentBgColor
    End If
    .CommentColor.Caption = ""
    
    'コメント フォント-------------------------------------------------------------------------------
'    Dim cBox As CommandBarComboBox
'    CommentFont = Library.getRegistry("Main", "CommentFont")
'
'    Set cBox = Application.CommandBars("Formatting").Controls.Item(1)
'
'      For i = 1 To cBox.ListCount
'        .CommentFont.AddItem cBox.list(i)
'        If cBox.list(i) = CommentFont Then
'          ListIndex = i - 1
'        End If
'      Next
'    .CommentFont.ListIndex = ListIndex

    
    imageName = thisAppName & "CommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoCommentImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    End If
    .CommentImg.Picture = LoadPicture(previewImgPath)
    
    '電子印鑑 フォント-------------------------------------------------------------------------------
    Dim cBox As CommandBarComboBox
    StampVal = Library.getRegistry("Main", "StampVal")
    StampFont = Library.getRegistry("Main", "StampFont")
  
    .StampVal.Value = StampVal
    Set cBox = Application.CommandBars("Formatting").Controls.Item(1)
      
      For i = 1 To cBox.ListCount
        .StampFont.AddItem cBox.list(i)
        If cBox.list(i) = StampFont Then
          ListIndex = i - 1
        End If
      Next
    .StampFont.ListIndex = ListIndex
    
    imageName = thisAppName & "StampImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoStampImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    End If
    .StampImg.Picture = LoadPicture(previewImgPath)
    
    
    
    
    
    
    
    
    '電子印鑑非表示(公開一時停止)
    .MultiPage1.Page4.Visible = False
    
  End With
  
  InitializeFlg = False
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
' * プレビュー表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function doHighLightPreview()
  Dim previewImgPath As String
  Dim HighLightColor As String, HighLightDspDirection As String, HighLightDspMethod As String, HighlightTransparentRate   As Long
  
  Call init.setting
'  Set BK_sheetHighLight = ActiveWorkbook.Worksheets("HighLight")
  
  
  HighLightColor = Me.HighLightColor.BackColor

  '透明度----------------------------------------------------------------------------------------
  HighlightTransparentRate = Me.HighlightTransparentRate.Value

  '表示方向--------------------------------------------------------------------------------------
  If Me.HighlightDspDirection_X.Value = True Then
    HighLightDspDirection = "X"
    
  ElseIf Me.HighlightDspDirection_Y.Value = True Then
    HighLightDspDirection = "Y"
    
  ElseIf Me.HighlightDspDirection_B.Value = True Then
    HighLightDspDirection = "B"
  End If
  
  '表示方法--------------------------------------------------------------------------------------
  If Me.HighlightDspMethod_0.Value = True Then
    HighLightDspMethod = "0"
  
  ElseIf Me.HighlightDspMethod_1.Value = True Then
    HighLightDspMethod = "1"
  
  ElseIf Me.HighlightDspMethod_2.Value = True Then
    HighLightDspMethod = "2"
  End If


  BK_sheetHighLight.Activate
  
  If BKh_rbPressed = False Then
    BKh_rbPressed = True
  End If
  
  Call Ctl_HighLight.showStart(Range("B2"), HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate)
  
  imageName = thisAppName & "HighLightImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(BK_sheetHighLight.Range("A1:C3"), imageName)
  
  
  If Library.chkFileExists(previewImgPath) = False Then
    
    imageName = thisAppName & "NoHighLightImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
  End If
  HighLightImg.Picture = LoadPicture(previewImgPath)
  
  BKh_rbPressed = old_BKh_rbPressed
  Call Ctl_HighLight.showStart(Range("C4"))
  
End Function


'==================================================================================================
Function doCommentPreview()
  Dim previewImgPath As String

  Call init.setting
'  Set BK_sheetHighLight = ActiveWorkbook.Worksheets("HighLight")
  
  BK_sheetHighLight.Activate
  BK_sheetHighLight.Range("N5").Activate
  
  CommentBgColor = Me.CommentColor.BackColor
  CommentFont = Me.StampFont.Value
  
  Call Library.setComment(CommentBgColor, CommentFont)
  
  imageName = thisAppName & "CommentImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(BK_sheetHighLight.Range("N4:S8"), imageName)
  
  
  If Library.chkFileExists(previewImgPath) = False Then
    imageName = thisAppName & "NoCommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
  End If
    CommentImg.Picture = LoadPicture(previewImgPath)
  
End Function


'==================================================================================================
Function doStampPreview()
  Dim previewImgPath As String
  Dim StampVal As String, StampFont As String
  

  Call init.setting(True)
  Set BK_sheetHighLight = ActiveWorkbook.Worksheets("HighLight")
  
  BK_sheetHighLight.Activate
  BK_sheetHighLight.Range("F10").Activate
  
  StampVal = Me.StampVal.Value
  StampFont = Me.StampFont.Value
  
  Call Ctl_Stamp.押印_確認印(StampVal, StampFont, thisAppName & "StampImg")
  
  imageName = thisAppName & "StampImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(BK_sheetHighLight.Range("E10:H12"), imageName)
  
  
  If Library.chkFileExists(previewImgPath) = False Then
    imageName = thisAppName & "NoCommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
  End If
  StampImg.Picture = LoadPicture(previewImgPath)
  
  BK_sheetHighLight.Shapes.Range(Array(thisAppName & "StampImg")).delete
  
  
End Function







'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'プレビュー
'==================================================================================================
Private Sub HighlightDspDirection_B_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightDspDirection_X_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightDspDirection_Y_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightDspMethod_0_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightDspMethod_1_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightDspMethod_2_Click()
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub

'==================================================================================================
Private Sub HighlightTransparentRate_Change()
  If InitializeFlg = False Then
    Call doHighLightPreview
    
    Me.HighlightTransparentRate_text.Caption = Me.HighlightTransparentRate.Value
  End If

End Sub


'==================================================================================================
Private Sub HighLightColor_Click()
  
  colorValue = Library.getColor(Me.HighLightColor.BackColor)
  Me.HighLightColor.BackColor = colorValue
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub


'==================================================================================================
Private Sub LineColor_Click()
  colorValue = Library.getColor(Me.LineColor.BackColor)
  Me.LineColor.BackColor = colorValue
  Me.LineColor.Caption = ""
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub


'==================================================================================================
Private Sub CommentColor_Click()
  colorValue = Library.getColor(Me.CommentColor.BackColor)
  Me.CommentColor.BackColor = colorValue
  Me.CommentColor.Caption = ""
  
  If InitializeFlg = False Then
    Call doCommentPreview
  End If
End Sub

'==================================================================================================
Private Sub StampFont_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  
  If InitializeFlg = False Then
    Call doStampPreview
  End If

End Sub

'==================================================================================================
Private Sub StampFont_Change()
  
  If InitializeFlg = False Then
    Call doStampPreview
  End If

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
  
  '透明度----------------------------------------------------------------------------------------
  Call Library.setRegistry("Main", "HighlightTransparentRate", HighlightTransparentRate.Value)

  '表示方向--------------------------------------------------------------------------------------
  If HighlightDspDirection_X.Value = True Then
    HighLightDspDirection = "X"
    
  ElseIf HighlightDspDirection_Y.Value = True Then
    HighLightDspDirection = "Y"
    
  ElseIf HighlightDspDirection_B.Value = True Then
    HighLightDspDirection = "B"
  End If
  Call Library.setRegistry("Main", "HighLightDspDirection", HighLightDspDirection)
  
  '表示方法--------------------------------------------------------------------------------------
  If HighlightDspMethod_0.Value = True Then
    HighLightDspMethod = "0"
  
  ElseIf HighlightDspMethod_1.Value = True Then
    HighLightDspMethod = "1"
  
  ElseIf HighlightDspMethod_2.Value = True Then
    HighLightDspMethod = "2"
  End If
  Call Library.setRegistry("Main", "HighLightDspMethod", HighLightDspMethod)
  BKh_rbPressed = old_BKh_rbPressed


  Call Library.setRegistry("Main", "CommentBgColor", Me.CommentColor.BackColor)


  Call Library.setRegistry("Main", "StampVal", Me.StampVal.Value)
  Call Library.setRegistry("Main", "StampFont", Me.StampFont.Value)
  
  
  'スタイルシートをスタイル2シートへコピー
'  endLine = sheetStyle2.Cells(Rows.count, 2).End(xlUp).Row
'  sheetStyle2.Range("A1:J" & endLine).Copy Destination:=sheetStyle.Range("A1")


  Unload Me
End Sub

  
