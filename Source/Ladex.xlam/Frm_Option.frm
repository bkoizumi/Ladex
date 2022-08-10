VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "�I�v�V����"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "Frm_Option.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
Public selectLine   As Long


Private Sub MultiPage1_Change()

End Sub

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim zoomLevelVal  As Variant
  Dim setZoomLevel As String
  Dim line As Long, endLine As Long
  Dim indexCnt As Integer, i As Variant
  Dim previewImgPath As String
  Dim ShortcutKeyList() As Variant
  Dim gridLineVal
  
  Dim cBox As CommandBarComboBox
  Const funcName As String = "Frm_Option.UserForm_Initialize"

  '�����J�n--------------------------------------
'  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  Application.Cursor = xlDefault
  InitializeFlg = True
  indexCnt = 0
  old_BKh_rbPressed = BKh_rbPressed
  
  '�\���ʒu�w��----------------------------------
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Width) / 2)
    
  setZoomLevel = Library.getRegistry("Main", "ZoomLevel")
  
  With Frm_Option
    .Caption = "�I�v�V���� |  " & thisAppName
    
    '��{�^�u-----------------------------------
    .userName.Value = Application.userName
    For Each zoomLevelVal In Split("25,50,75,85,100", ",")
      .ZoomLevel.AddItem zoomLevelVal
      If zoomLevelVal = setZoomLevel Then
        .ZoomLevel.ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    
    '�g���̕\��----------------------------------
    .GridLine.AddItem "�\�����Ȃ�"
    .GridLine.AddItem "�\������"
    .GridLine.AddItem "�ύX���Ȃ�"
    
    gridLineVal = Library.getRegistry("Main", "gridLine")
    If gridLineVal = False Then
      .GridLine.Text = "�\�����Ȃ�"
    ElseIf gridLineVal = True Then
      .GridLine.Text = "�\������"
    Else
      .GridLine.Text = gridLineVal
    End If
    
    '�s�̍����A��̕�----------------------------
    .ColumnWidth.Value = Library.getRegistry("Main", "ColumnWidth")
    .rowHeight.Value = Library.getRegistry("Main", "rowHeight")
    
    .BgColor.Value = Library.getRegistry("Main", "bgColor")
      
    LineColor = Library.getRegistry("Main", "LineColor")
    If LineColor = "" Then
      .LineColor.BackColor = 0
    Else
      .LineColor.BackColor = LineColor
    End If
    .LineColor.Caption = ""
  
    '��{�t�H���g--------------------------------
    BaseFont = Library.getRegistry("Main", "BaseFont")
    Set cBox = Application.CommandBars("Formatting").Controls.Item(1)
    indexCnt = 0
    For i = 1 To cBox.ListCount
      .BaseFont.AddItem cBox.list(i)
      If cBox.list(i) = BaseFont Then
        ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .BaseFont.ListIndex = ListIndex

    '��{�t�H���g�T�C�Y
    indexCnt = 0
    BaseFontSize = Library.getRegistry("Main", "BaseFontSize")
    For Each i In Split("6,7,8,9,10,11,12,14,16,18,20", ",")
      .BaseFontSize.AddItem i
      If i = BaseFontSize Then
        ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .BaseFontSize.ListIndex = ListIndex
  
    '�f�o�b�O�֘A--------------------------------
    .LogLevel.AddItem "1.Error"
    .LogLevel.AddItem "2.warning"
    .LogLevel.AddItem "3.notice"
    .LogLevel.AddItem "4.info"
    .LogLevel.AddItem "5.debug"
    .LogLevel.Text = BK_setVal("LogLevel")
    
    .debugMode.AddItem "all"
    .debugMode.AddItem "File"
    .debugMode.AddItem "Speak"
    .debugMode.AddItem "none"
    .debugMode.AddItem "develop"
    .debugMode.Text = BK_setVal("debugMode")
  
    '�n�C���C�g�^�u------------------------------
    HighLightColor = Library.getRegistry("Main", "HighLightColor")
    If HighLightColor = "" Then
      .HighLightColor.BackColor = 10222585
    Else
      .HighLightColor.BackColor = HighLightColor
    End If
    .HighLightColor.Caption = ""
    
    '�����x
    .HighlightTransparentRate.Min = 0
    .HighlightTransparentRate.max = 100
    HighlightTransparentRate = Library.getRegistry("Main", "HighLightTransparentRate")
    If HighlightTransparentRate = "0" Then
      .HighlightTransparentRate.Value = 70
      .HighlightTransparentRate_text.Caption = 70
    Else
      .HighlightTransparentRate.Value = HighlightTransparentRate
      .HighlightTransparentRate_text.Caption = HighlightTransparentRate
    End If
  
    '�\������
    HighLightDspDirection = Library.getRegistry("Main", "HighLightDspDirection")
    If HighLightDspDirection = "X" Then
      .HighlightDspDirection_X.Value = True
      
    ElseIf HighLightDspDirection = "Y" Then
      .HighlightDspDirection_Y.Value = True
    
    ElseIf HighLightDspDirection = "B" Then
      .HighlightDspDirection_B.Value = True
    
    End If
  
    '�\�����@
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
    
    '�v���r���[
    imageName = thisAppName & "HighLightImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoHighLightImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    Else
      Call doHighLightPreview
    End If
    HighLightImg.Picture = LoadPicture(previewImgPath)
    
    
    '�R�����g�^�u--------------------------------
    commentBgColor = Library.getRegistry("Main", "CommentBgColor")
    .CommentColor.BackColor = commentBgColor
    .CommentColor.Caption = ""
    
    '�R�����g �t�H���g
    CommentFontColor = Library.getRegistry("Main", "CommentFontColor")
    .CommentFontColor.BackColor = CommentFontColor
    .CommentFontColor.Caption = ""
    
    CommentFont = Library.getRegistry("Main", "CommentFont")
    Set cBox = Application.CommandBars("Formatting").Controls.Item(1)
    indexCnt = 0
    For i = 1 To cBox.ListCount
      .CommentFont.AddItem cBox.list(i)
      If cBox.list(i) = CommentFont Then
        ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .CommentFont.ListIndex = ListIndex

    '�R�����g �t�H���g�T�C�Y
    indexCnt = 0
    CommentFontSize = Library.getRegistry("Main", "CommentFontSize")
    For Each i In Split("6,7,8,9,10,11,12,14,16,18,20", ",")
      .CommentFontSize.AddItem i
      If i = CommentFontSize Then
        ListIndex = indexCnt
      End If
      indexCnt = indexCnt + 1
    Next
    .CommentFontSize.ListIndex = ListIndex

    '�R�����g �v���r���[
    imageName = thisAppName & "CommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoCommentImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    Else
      Call doCommentPreview
    End If
    .CommentImg.Picture = LoadPicture(previewImgPath)
    Set cBox = Nothing
    
    '�d�q��Ӄ^�u--------------------------------
    '.MultiPage1.Page4.Visible = False        '�d�q��Ӕ�\��(���J�ꎞ��~)

    StampName = Library.getRegistry("Main", "StampName")
    StampFont = Library.getRegistry("Main", "StampFont")
    StampVal = Library.getRegistry("Main", "StampVal")

    .StampName.Value = StampName
    .StampVal.Value = StampVal
    
    Set cBox = Application.CommandBars("Formatting").Controls.Item(1)
      For i = 1 To cBox.ListCount
        .StampFont.AddItem cBox.list(i)
        If cBox.list(i) = StampFont Then
          ListIndex = i - 1
        End If
      Next
    .StampFont.ListIndex = ListIndex



    '�v���r���[
    imageName = thisAppName & "StampImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
    If Library.chkFileExists(previewImgPath) = False Then
      imageName = thisAppName & "NoStampImg" & ".jpg"
      previewImgPath = LadexDir & "\RibbonImg\" & imageName
    End If
    .StampImg.Picture = LoadPicture(previewImgPath)
    Set cBox = Nothing
    
    '�V���[�g�J�b�g�^�u-------------------------
    onAlt.Value = True
    indexCnt = 1
    With funcList
      .View = lvwReport
      .LabelEdit = lvwManual
      .HideSelection = False
      .AllowColumnReorder = True
      .FullRowSelect = True
      .Gridlines = True
      .ColumnHeaders.add , "_ID", "#", 30
      .ColumnHeaders.add , "_ShortcutKey", "�L�["
      .ColumnHeaders.add , "_Label", "�@�\����", 100
      .ColumnHeaders.add , "_description", "�@�\", 400
      .ColumnHeaders.add , "_KeyID", "KeyID", 0
      
      endLine = LadexSh_Function.Cells(Rows.count, 1).End(xlUp).Row
      For line = 2 To endLine
        If LadexSh_Function.Range("D" & line).Value <> "" Then
          With .ListItems.add
            .Text = indexCnt
            .SubItems(1) = LadexSh_Function.Range("B" & line).Value
            .SubItems(2) = LadexSh_Function.Range("C" & line).Value
            .SubItems(3) = LadexSh_Function.Range("D" & line).Value
            .SubItems(4) = LadexSh_Function.Range("F" & line).Value
          End With
          indexCnt = indexCnt + 1
        End If
      Next
    End With
    
    '�L�[���X�g
    endLine = LadexSh_Config.Cells(Rows.count, 13).End(xlUp).Row
    endLine = 59
    
    ReDim ShortcutKeyList(endLine - 3, 2)
    For line = 3 To endLine
      If LadexSh_Config.Range("N" & line) <> "" Then
        ShortcutKeyList(line - 3, 0) = CStr(LadexSh_Config.Range("N" & line))
        ShortcutKeyList(line - 3, 1) = CStr(LadexSh_Config.Range("M" & line))
        ShortcutKeyList(line - 3, 2) = CStr(LadexSh_Config.Range("O" & line))
      End If
    Next
    With ShortcutKey
      .ColumnCount = 3
      .TextColumn = 1
      .BoundColumn = 1
      .ColumnWidths = "60;0;0"
      .list() = ShortcutKeyList()
    End With

    
    
    
    .MultiPage1.Value = 0
  End With
  
  InitializeFlg = False
  Exit Sub
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Sub


'**************************************************************************************************
' * �X�^�C���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub IncludeFont01_Click()
  
  If IncludeFont01.Value = True Then
    ret = �Z���̏����ݒ�_�t�H���g(1)
    IncludeFont01.Value = ret
  End If
End Sub

'**************************************************************************************************
' * �g�ݍ��݃_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �Z���̏����ݒ�_�t�H���g(Optional line As Long = 1)
  Call init.setting
  sheetStyle2.Select
  sheetStyle2.Cells(line + 1, 11).Select
  ret = Application.Dialogs(xlDialogActiveCellFont).Show
  If ret = True Then
    sheetStyle2.Cells(line + 1, 5) = "TRUE"
  Else
    sheetStyle2.Cells(line + 1, 5) = "FALSE"
  End If
  �Z���̏����ݒ�_�t�H���g = ret
End Function


'**************************************************************************************************
' * �v���r���[�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function doHighLightPreview()
  Dim previewImgPath As String
  Dim HighLightColor As String, HighLightDspDirection As String, HighLightDspMethod As String, HighlightTransparentRate   As Long
  
  Call init.setting

  
  HighLightColor = Me.HighLightColor.BackColor

  '�����x----------------------------------------------------------------------------------------
  HighlightTransparentRate = Me.HighlightTransparentRate.Value

  '�\������--------------------------------------------------------------------------------------
  If Me.HighlightDspDirection_X.Value = True Then
    HighLightDspDirection = "X"
    
  ElseIf Me.HighlightDspDirection_Y.Value = True Then
    HighLightDspDirection = "Y"
    
  ElseIf Me.HighlightDspDirection_B.Value = True Then
    HighLightDspDirection = "B"
  End If
  
  '�\�����@--------------------------------------------------------------------------------------
  If Me.HighlightDspMethod_0.Value = True Then
    HighLightDspMethod = "0"
  
  ElseIf Me.HighlightDspMethod_1.Value = True Then
    HighLightDspMethod = "1"
  
  ElseIf Me.HighlightDspMethod_2.Value = True Then
    HighLightDspMethod = "2"
  End If
  
  LadexSh_HiLight.Activate
  
  If BKh_rbPressed = False Then
    BKh_rbPressed = True
  End If
'  Range("A1:D4").Clear
'  Call Library.�r��_����_�i�q(Range("A1:C3"))
  
  'Call Ctl_HighLight.showStart(Range("B2"), HighLightColor, HighLightDspDirection, HighLightDspMethod, HighlightTransparentRate)
  Call Ctl_HighLight.showStart(Range("B2"))
  
  imageName = thisAppName & "HighLightImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(LadexSh_HiLight.Range("A1:C3"), imageName)
  
  
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
  Dim commentBgColor, CommentFontColor, CommentFont, CommentFontSize

  Call init.setting
'  Set LadexSh_HiLight = ActiveWorkbook.Worksheets("HighLight")
  
  LadexSh_HiLight.Activate
  LadexSh_HiLight.Range("N7").Activate
  
  commentBgColor = Me.CommentColor.BackColor
  CommentFontColor = Me.CommentFontColor.BackColor
  CommentFont = Me.CommentFont.Value
  CommentFontSize = Me.CommentFontSize.Value
  
  Call Library.setComment(commentBgColor, CommentFont, CommentFontColor, CommentFontSize)
  
  imageName = thisAppName & "CommentImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(LadexSh_HiLight.Range("N6:R9"), imageName)
  
  If Library.chkFileExists(previewImgPath) = False Then
    imageName = thisAppName & "NoCommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
  End If
    CommentImg.Picture = LoadPicture(previewImgPath)
End Function


'==================================================================================================
Function doStampPreview()
  Dim previewImgPath As String
  
  Call init.setting(True)
  
  LadexSh_HiLight.Activate
  LadexSh_HiLight.Range("F12").Activate
  
  Call Ctl_Stamp.�m�F��(StampName.Value, StampVal.Value, StampFont.Value, thisAppName & "StampImg")
  
  imageName = thisAppName & "StampImg" & ".jpg"
  previewImgPath = LadexDir & "\RibbonImg\" & imageName
  Call Ctl_Image.saveSelectArea2Image(LadexSh_HiLight.Range("E12:I16"), imageName)
  
  
  If Library.chkFileExists(previewImgPath) = False Then
    imageName = thisAppName & "NoCommentImg" & ".jpg"
    previewImgPath = LadexDir & "\RibbonImg\" & imageName
  End If
  StampImg.Picture = LoadPicture(previewImgPath)
  
  LadexSh_HiLight.Shapes.Range(Array(thisAppName & "StampImg")).delete
End Function


'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�v���r���[
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
    
    HighlightTransparentRate_text.Caption = HighlightTransparentRate.Value
  End If

End Sub


'==================================================================================================
Private Sub HighLightColor_Click()
  
  colorValue = Library.getColor(HighLightColor.BackColor)
  HighLightColor.BackColor = colorValue
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub


'==================================================================================================
Private Sub LineColor_Click()
  colorValue = Library.getColor(LineColor.BackColor)
  LineColor.BackColor = colorValue
  LineColor.Caption = ""
  
  If InitializeFlg = False Then
    Call doHighLightPreview
  End If
End Sub


'==================================================================================================
Private Sub CommentColor_Click()
  colorValue = Library.getColor(CommentColor.BackColor)
  CommentColor.BackColor = colorValue
  CommentColor.Caption = ""
  
  If InitializeFlg = False Then
    Call doCommentPreview
  End If
End Sub

'==================================================================================================
Private Sub CommentFontColor_Click()
  colorValue = Library.getColor(CommentFontColor.BackColor)
  CommentFontColor.BackColor = colorValue
  CommentFontColor.Caption = ""
  
  If InitializeFlg = False Then
    Call doCommentPreview
  End If
  CommentFontColor.Caption = ""
End Sub

'==================================================================================================
Private Sub CommentFont_Change()
  CommentFontColor.Caption = ""
  If InitializeFlg = False Then
    Call doCommentPreview
  End If
End Sub

'==================================================================================================
Private Sub CommentFontSize_Change()
  CommentFontColor.Caption = ""
  If InitializeFlg = False Then
    Call doCommentPreview
  End If
End Sub

'==================================================================================================
Private Sub StampName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  If InitializeFlg = False Then
    Call doStampPreview
  End If
End Sub

'==================================================================================================
Private Sub StampName_Change()
  If InitializeFlg = False Then
    Call doStampPreview
  End If
End Sub

'==================================================================================================
Private Sub StampVal_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  If InitializeFlg = False Then
    Call doStampPreview
  End If
End Sub

'==================================================================================================
Private Sub StampVal_Change()
  If InitializeFlg = False Then
    Call doStampPreview
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

''==================================================================================================
Private Sub setShortcutKey_Click()
  Dim keyVal As String
  
  Call init.setting
  
  Call Library.showDebugForm("funcList.Item    ", funcList.SelectedItem.Text, "debug")
  Call Library.showDebugForm("funcList.SubItem1", funcList.SelectedItem.SubItems(1), "debug")
  Call Library.showDebugForm("funcList.SubItem2", funcList.SelectedItem.SubItems(2), "debug")
  Call Library.showDebugForm("funcList.SubItem3", funcList.SelectedItem.SubItems(3), "debug")

  Call Library.showDebugForm("onCtrl ", onCtrl.Value, "debug")
  Call Library.showDebugForm("onAlt  ", onAlt.Value, "debug")
  Call Library.showDebugForm("onShift", onShift.Value, "debug")
  Call Library.showDebugForm("ShortcutKey", ShortcutKey.list(ShortcutKey.ListIndex, 0), "debug")
  Call Library.showDebugForm("ShortcutKey", ShortcutKey.list(ShortcutKey.ListIndex, 1), "debug")
  Call Library.showDebugForm("ShortcutKey", ShortcutKey.list(ShortcutKey.ListIndex, 2), "debug")

  selectLine = funcList.SelectedItem.Text

  LadexSh_Function.Range("E" & selectLine + 1, "F" & selectLine + 1) = ""

  keyVal = ""
  If onCtrl.Value = True Then
    keyVal = "Ctrl"
  End If
  
  If onAlt.Value = True Then
    If keyVal = "" Then
      keyVal = "Alt"
    Else
      keyVal = keyVal & "+Alt"
    End If
  End If
  
  If onShift.Value = True Then
    If keyVal = "" Then
      keyVal = "Shift"
    Else
      keyVal = keyVal & "+Shift"
    End If
  End If
  
  If keyVal = "" Then
    keyVal = "Alt"
  End If
'  keyVal = keyVal & "+" & ShortcutKey.list(ShortcutKey.ListIndex, 1)
  
  Call Library.showDebugForm("keyVal", keyVal, "debug")
  If WorksheetFunction.CountIf(LadexSh_Function.Range("B2:B1000"), keyVal & "+" & ShortcutKey.list(ShortcutKey.ListIndex, 0)) > 1 Then
    megLabel.Caption = "�����ݒ肪���łɂ���܂�"
  Else
    LadexSh_Function.Range("B" & selectLine + 1) = keyVal & "+" & ShortcutKey.list(ShortcutKey.ListIndex, 0)
    LadexSh_Function.Range("E" & selectLine + 1) = keyVal & "+" & ShortcutKey.list(ShortcutKey.ListIndex, 2)
    LadexSh_Function.Range("F" & selectLine + 1) = ShortcutKey.list(ShortcutKey.ListIndex, 1)
  
    Call reLoadFuncList
  End If
End Sub

'==================================================================================================
Function reLoadFuncList()

  funcList.ListItems.Clear
  funcList.ColumnHeaders.Clear
  With funcList
    .View = lvwReport
    .LabelEdit = lvwManual
    .HideSelection = False
    .AllowColumnReorder = True
    .FullRowSelect = True
    .Gridlines = True
    .ColumnHeaders.add , "_ID", "#", 30
    .ColumnHeaders.add , "_ShortcutKey", "�L�["
    .ColumnHeaders.add , "_Label", "�@�\����", 100
    .ColumnHeaders.add , "_description", "�@�\", 400
    .ColumnHeaders.add , "_KeyID", "KeyID", 0
    
    endLine = LadexSh_Function.Cells(Rows.count, 1).End(xlUp).Row
    For line = 2 To endLine
      If LadexSh_Function.Range("C" & line).Value <> "" Then
        With .ListItems.add
          .Text = LadexSh_Function.Range("A" & line).Value
          .SubItems(1) = LadexSh_Function.Range("B" & line).Value
          .SubItems(2) = LadexSh_Function.Range("C" & line).Value
          .SubItems(3) = LadexSh_Function.Range("D" & line).Value
          .SubItems(4) = LadexSh_Function.Range("F" & line).Value
        End With
      End If
    Next
    .ListItems(selectLine).EnsureVisible
    .ListItems(selectLine).Selected = True
    .SetFocus

  End With
End Function

'==================================================================================================
Private Sub funcList_Click()
  Dim keyVal As Variant
  
  Call Library.showDebugForm("funcList.Item    ", funcList.SelectedItem.Text, "debug")
  Call Library.showDebugForm("funcList.SubItem1", funcList.SelectedItem.SubItems(1), "debug")
  Call Library.showDebugForm("funcList.SubItem2", funcList.SelectedItem.SubItems(2), "debug")
  Call Library.showDebugForm("funcList.SubItem3", funcList.SelectedItem.SubItems(3), "debug")
  Call Library.showDebugForm("funcList.SubItem4", funcList.SelectedItem.SubItems(4), "debug")
  
  If funcList.SelectedItem.SubItems(1) <> "" Then
    onCtrl.Value = False
    onAlt.Value = False
    onShift.Value = False
    For Each keyVal In Split(funcList.SelectedItem.SubItems(1), "+")
      If keyVal = "Ctrl" Then
        onCtrl.Value = True
      ElseIf keyVal = "Alt" Then
        onAlt.Value = True
      ElseIf keyVal = "Shift" Then
        onShift.Value = True
      Else
        ShortcutKey.Value = keyVal
      End If
    Next
  Else
    onCtrl.Value = False
    onAlt.Value = True
    onShift.Value = False
    ShortcutKey.Value = ""
  End If
End Sub

'==================================================================================================
Private Sub Del_ShortcutKey_Click()
  
  Call init.setting
  selectLine = funcList.SelectedItem.Text
  
  LadexSh_Function.Range("B" & selectLine + 1) = ""
  LadexSh_Function.Range("E" & selectLine + 1) = ""
  LadexSh_Function.Range("F" & selectLine + 1) = ""
  megLabel.Caption = ""
  Call reLoadFuncList
End Sub






'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
'  Call Library.setRegistry("UserForm", "OptionTop", Top)
'  Call Library.setRegistry("UserForm", "OptionLeft", Left)
  
  InitializeFlg = True
  
  Unload Me
End Sub

'==================================================================================================
' ���s
Private Sub run_Click()
  Dim execDay As Date
  
  InitializeFlg = True
  
'  Call Library.setRegistry("UserForm", "OptionTop", Top)
'  Call Library.setRegistry("UserForm", "OptionLeft", Left)
  
  Call Library.setRegistry("Main", "ZoomLevel", ZoomLevel.Text)
  Call Library.setRegistry("Main", "GridLine", GridLine.Value)
  Call Library.setRegistry("Main", "bgColor", BgColor.Value)
  Call Library.setRegistry("Main", "LineColor", LineColor.BackColor)
  Call Library.setRegistry("Main", "debugMode", debugMode.Value)
  Call Library.setRegistry("Main", "LogLevel", LogLevel.Value)
  
  Call Library.setRegistry("Main", "BaseFont", BaseFont.Value)
  Call Library.setRegistry("Main", "BaseFontSize", BaseFontSize.Value)
  Call Library.setRegistry("Main", "rowHeight", rowHeight.Value)
  Call Library.setRegistry("Main", "ColumnWidth", ColumnWidth.Value)
  Call Library.setRegistry("Main", "LogLevel", LogLevel.Value)
  Call Library.setRegistry("Main", "LogLevel", LogLevel.Value)
  
  
  Call Library.setRegistry("Main", "debugMode", debugMode.Value)
  Call Library.setRegistry("Main", "LogLevel", LogLevel.Value)
  
  BK_setVal("debugMode") = debugMode.Value
  BK_setVal("LogLevel") = LogLevel.Value
  
  Application.userName = userName.Value
  
  '�n�C���C�g�ݒ�--------------------------------
  Call Library.setRegistry("Main", "HighLightColor", HighLightColor.BackColor)
  
  '�����x
  Call Library.setRegistry("Main", "HighlightTransparentRate", HighlightTransparentRate.Value)

  '�\������
  If HighlightDspDirection_X.Value = True Then
    HighLightDspDirection = "X"
    
  ElseIf HighlightDspDirection_Y.Value = True Then
    HighLightDspDirection = "Y"
    
  ElseIf HighlightDspDirection_B.Value = True Then
    HighLightDspDirection = "B"
  End If
  Call Library.setRegistry("Main", "HighLightDspDirection", HighLightDspDirection)
  
  '�\�����@
  If HighlightDspMethod_0.Value = True Then
    HighLightDspMethod = "0"
  
  ElseIf HighlightDspMethod_1.Value = True Then
    HighLightDspMethod = "1"
  
  ElseIf HighlightDspMethod_2.Value = True Then
    HighLightDspMethod = "2"
  End If
  Call Library.setRegistry("Main", "HighLightDspMethod", HighLightDspMethod)
  BKh_rbPressed = old_BKh_rbPressed

  '�R�����g�ݒ�----------------------------------
  Call Library.setRegistry("Main", "CommentBgColor", CommentColor.BackColor)
  Call Library.setRegistry("Main", "CommentFont", CommentFont.Value)
  
  Call Library.setRegistry("Main", "CommentFontColor", CommentFontColor.BackColor)
  Call Library.setRegistry("Main", "CommentFontSize", CommentFontSize.Value)

  '�X�^���v--------------------------------------
  Call Library.setRegistry("Main", "StampName", StampName.Value)
  Call Library.setRegistry("Main", "StampVal", StampVal.Value)
  Call Library.setRegistry("Main", "StampFont", StampFont.Value)
  
  '�X�^�C���V�[�g���X�^�C��2�V�[�g�փR�s�[
'  endLine = sheetStyle2.Cells(Rows.count, 2).End(xlUp).Row
'  sheetStyle2.Range("A1:J" & endLine).Copy Destination:=sheetStyle.Range("A1")
  
  Unload Me
End Sub

