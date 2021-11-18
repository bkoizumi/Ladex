Attribute VB_Name = "Ctl_Ribbon_old"
Option Explicit

Private ctlEvent As New clsEvent

#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If


'**************************************************************************************************
' * ���{�����j���[�����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�ǂݍ��ݎ�����
Function onLoad(ribbon As IRibbonUI)
  
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  
  
  Set BK_ribbonUI = ribbon
  
  BKh_rbPressed = Library.getRegistry("Main", "HighLightFlg")
  BKz_rbPressed = Library.getRegistry("Main", "ZoomFlg")
  BKT_rbPressed = Library.getRegistry("Main", "CustomRibbon")
  
  Call Library.setRegistry("Main", "BK_ribbonUI", CStr(ObjPtr(BK_ribbonUI)))
  
  BK_ribbonUI.ActivateTab ("Ladex")
  BK_ribbonUI.Invalidate
  
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
  Call Library.errorHandle
End Function


'==================================================================================================
' �g�O���{�^���Ƀ`�F�b�N��ݒ肷��
Function HighLightPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKh_rbPressed
End Function

'==================================================================================================
' �g�O���{�^���Ƀ`�F�b�N��ݒ肷��
Function ZoomInPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKz_rbPressed
End Function

'==================================================================================================
' �g�O���{�^���Ƀ`�F�b�N��ݒ肷��
Function confFormulaPressed(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = BKcf_rbPressed
End Function

'==================================================================================================
Function noDispTab(control As IRibbonControl)
  Call Library.setRegistry("Main", "CustomRibbon", False)
  Call RefreshRibbon
End Function

'==================================================================================================
Function setDispTab(control As IRibbonControl, pressed As Boolean)
  
  BKT_rbPressed = pressed
  Call Library.setRegistry("Main", "CustomRibbon", pressed)
  Call RefreshRibbon
End Function

'==================================================================================================
Function getDispTab(control As IRibbonControl, ByRef returnedVal)
  
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
End Function

 
  
'==================================================================================================
'�V�[�g�ꗗ���j���[
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa, sheetNameID
  
  On Error GoTo catchError
  
  Call init.setting
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
      With MenuSepa
        .SetAttribute "id", "M_RelaxTools"
        .SetAttribute "title", "RelaxTools�𗘗p"
      End With
      Menu.AppendChild MenuSepa
      Set MenuSepa = Nothing

      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "RelaxTools"
        .SetAttribute "label", "RelaxTools"
        .SetAttribute "supertip", "RelaxTools�̃V�[�g�Ǘ����N��"
        
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
        .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.actRelaxSheetManager"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
  End If
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "sheetID_0"
      .SetAttribute "title", ActiveWorkbook.Name
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
  
  
  
  For Each sheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = sheetName.Name
      .SetAttribute "id", "sheetID_" & sheetName.Index
      .SetAttribute "label", sheetName.Name
    
      If ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
        .SetAttribute "supertip", "�A�N�e�B�u�V�[�g"
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
        
      ElseIf Sheets(sheetName.Name).Visible = True Then
       '.SetAttribute "supertip", "�A�N�e�B�u�V�[�g"
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      
      ElseIf Sheets(sheetName.Name).Visible = 0 Then
        .SetAttribute "supertip", "��\���V�[�g"
        .SetAttribute "imageMso", "SheetProtect"
      
      ElseIf Sheets(sheetName.Name).Visible = 2 Then
        .SetAttribute "supertip", "�}�N���ɂ���\���V�[�g"
        .SetAttribute "imageMso", "ReviewProtectWorkbook"
      
      End If
      
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Next

  DOMDoc.AppendChild Menu
  
'  Debug.Print DOMDoc.XML
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  BK_ribbonUI.Invalidate

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function

'==================================================================================================
Function dMenuRefresh(control As IRibbonControl)
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  BK_ribbonUI.Invalidate
End Function

'==================================================================================================
'RelaxTools
Function getRelaxTools(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa
  
  
'  On Error GoTo catchError
  Call init.setting
  
    If BK_ribbonUI Is Nothing Then
      #If VBA7 And Win64 Then
        Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
      #Else
        Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
      #End If
    End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"
  
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    'RelaxTools�擾------------------------------------------------------------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxToolsGet"
      .SetAttribute "title", "RelaxTool�����"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools_get"
      .SetAttribute "label", "RelaxTool�����"
      .SetAttribute "image", "RelaxToolsLogo"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxTools------------------------------------------------------------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxTools"
      .SetAttribute "title", "RelaxTools�𗘗p"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools01"
      .SetAttribute "label", "�V�[�g�Ǘ�"
      .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.actRelaxSheetManager"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxShapes------------------------------------------------------------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxShapes"
      .SetAttribute "title", "RelaxShapes�𗘗p"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes01"
      .SetAttribute "label", "�T�C�Y���킹"
      .SetAttribute "imageMso", "ShapesDuplicate"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes01"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes02"
      .SetAttribute "label", "��ʒu���킹"
      .SetAttribute "imageMso", "ObjectsAlignTop"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes02"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxShapes03"
      .SetAttribute "label", "���ʒu���킹"
      .SetAttribute "imageMso", "ObjectsAlignLeft"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes03"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
    
    'RelaxApps------------------------------------------------------------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxApps"
      .SetAttribute "title", "RelaxApps�𗘗p"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxApps01"
      .SetAttribute "label", "�t�k�r��"
      .SetAttribute "imageMso", "BorderDrawGrid"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxApps01"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  Else
    'RelaxTools�擾------------------------------------------------------------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "M_RelaxTools"
      .SetAttribute "title", "RelaxTool�����"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "RelaxTools_get"
      .SetAttribute "label", "RelaxTool�����"
      .SetAttribute "image", "RelaxToolsLogo"
      .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  End If

  DOMDoc.AppendChild Menu
  
'  Call Library.showDebugForm(DOMDoc.XML)
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  'BK_ribbonUI.Invalidate
  BK_ribbonUI.InvalidateControl control.ID

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
'  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function selectActiveSheet(control As IRibbonControl)
  Dim sheetNameID As Integer
  Dim sheetCount As Integer
  Dim sheetName As Worksheet
  
  Call Library.startScript
  sheetNameID = Replace(control.ID, "sheetID_", "")
  
  If Sheets(sheetNameID).Visible <> 2 Then
    Sheets(sheetNameID).Visible = True
  
  ElseIf Sheets(sheetNameID).Visible = 2 Then
    If MsgBox("�}�N���ɂ���Ĕ�\���ƂȂ��Ă���V�[�g�ł�" & vbNewLine & "�}�N���̓���ɉe����^����\��������܂��B" & vbNewLine & "�\�����܂����H", vbYesNo + vbCritical) = vbNo Then
      End
    Else
      Sheets(sheetNameID).Visible = True
    End If
  End If
  
  sheetCount = 1
  For Each sheetName In ActiveWorkbook.Sheets
    If Sheets(sheetName.Name).Visible = True And sheetName.Name = Sheets(sheetNameID).Name Then
      Exit For
    Else
      sheetCount = sheetCount + 1
    End If
  Next
  ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
  ActiveWindow.ScrollWorkbookTabs Sheets:=sheetCount
  Sheets(sheetNameID).Select
  
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  
  If BK_ribbonUI Is Nothing Then
  Else
    BK_ribbonUI.Invalidate
  End If
  
  
  Call Library.endScript
End Function


'==================================================================================================
Function encode(strVal As String)

  strVal = Replace(strVal, "(", "BK1_")
  strVal = Replace(strVal, ")", "BK2_")
  strVal = Replace(strVal, " ", "BK3_")
  strVal = Replace(strVal, "�@", "BK4_")
  strVal = Replace(strVal, "�y", "BK5_")
  strVal = Replace(strVal, "�z", "BK6_")
  strVal = Replace(strVal, "�i", "BK7_")
  strVal = Replace(strVal, "�j", "BK8_")
  
  strVal = "BK0_" & strVal
  encode = strVal
End Function

'==================================================================================================
Function decode(strVal As String)

  strVal = Replace(strVal, "BK0_", "")
  strVal = Replace(strVal, "BK1_", "(")
  strVal = Replace(strVal, "BK2_", ")")
  strVal = Replace(strVal, "BK3_", " ")
  strVal = Replace(strVal, "BK4_", "�@")
  strVal = Replace(strVal, "BK5_", "�y")
  strVal = Replace(strVal, "BK6_", "�z")
  strVal = Replace(strVal, "BK7_", "�i")
  strVal = Replace(strVal, "BK8_", "�j")
  
  
  decode = strVal
End Function


'==================================================================================================
#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If

  Dim ribbonObj As Object
  
  If lRibbonPointer = 0 Then
    End
  End If
  
'  Stop
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p)
End Function





' ���C�ɓ��胁�j���[�쐬---------------------------------------------------------------------------
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa
  
  On Error GoTo catchError
  Call init.setting
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menu�̍쐬

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  Call Ctl_Favorite.getList
  If Workbooks.count = 0 Then
    endLine = 100
  Else
    endLine = BK_sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  End If
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .SetAttribute "id", "MS_���C�ɓ���ꗗ"
      .SetAttribute "title", "���C�ɓ���ꗗ"
    End With
    Menu.AppendChild MenuSepa
    Set MenuSepa = Nothing
    
    
  For line = 2 To endLine
    If BK_sheetFavorite.Range("A" & line) <> "" Then
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "Favorite_" & line
        .SetAttribute "label", objFso.GetFileName(BK_sheetFavorite.Range("A" & line))
        .SetAttribute "imageMso", "Favorites"
        .SetAttribute "supertip", BK_sheetFavorite.Range("A" & line)
        .SetAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.OpenFavoriteList"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
    End If
  Next
  DOMDoc.AppendChild Menu
  returnedVal = DOMDoc.XML
'  Call Library.showDebugForm(DOMDoc.XML)
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function OpenFavoriteList(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  
  line = Replace(control.ID, "Favorite_", "")
  fileNamePath = BK_sheetFavorite.Range("A" & line)
  
  If Library.chkFileExists(fileNamePath) Then
    Workbooks.Open fileName:=fileNamePath
  Else
    MsgBox "�t�@�C�������݂��܂���" & vbNewLine & fileNamePath, vbExclamation
  End If
End Function





'Label �ݒ�----------------------------------------------------------------------------------------
Public Function getLabel(control As IRibbonControl, ByRef setRibbonVal)
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = Replace(BK_ribbonVal("Lbl_" & control.ID), "<BR>", vbNewLine)
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function

'Action �ݒ�---------------------------------------------------------------------------------------
Function getAction(control As IRibbonControl)
  Dim setRibbonVal As Variant
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = BK_ribbonVal("Act_" & control.ID)
  
  If setRibbonVal Like "*Ctl_Ribbon.*" Then
    Call Application.run(setRibbonVal, control)
  
  ElseIf setRibbonVal = "" Then
    Call Library.showDebugForm("Act_" & control.ID)
  Else
    Call Application.run(setRibbonVal)
  End If
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
  Call Library.endScript
End Function


'Supertip �ݒ�-------------------------------------------------------------------------------------
Public Function getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  Call init.setting
  setRibbonVal = BK_ribbonVal("Sup_" & control.ID)
End Function


'Description �ݒ�----------------------------------------------------------------------------------
Public Function getDescription(control As IRibbonControl, ByRef setRibbonVal)
  Call init.setting
  setRibbonVal = Replace(BK_ribbonVal("Dec_" & control.ID), "<BR>", vbNewLine)

End Function

'getImageMso �ݒ�----------------------------------------------------------------------------------
Public Function getImage(control As IRibbonControl, ByRef image)
  On Error GoTo catchError
  
  Call init.setting
  image = BK_ribbonVal("Img_" & control.ID)
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function


'size �ݒ�-----------------------------------------------------------------------------------------
Public Function getSize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  
  On Error GoTo catchError
  
  Call init.setting
  setRibbonVal = BK_ribbonVal("Siz_" & control.ID)
  Select Case setRibbonVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
      setRibbonVal = 0
  End Select
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function

'==================================================================================================
'�L��/�����؂�ւ�
Function getEnabled(control As IRibbonControl, ByRef returnedVal)
  Dim wb As Workbook
  Call init.setting
  
  If Workbooks.count = 0 Then
    returnedVal = False
  ElseIf BK_setVal("debugMode") = "develop" Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'==================================================================================================
Function getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
  Call RefreshRibbon
    
End Function


'==================================================================================================
Function RefreshRibbon()
  On Error GoTo catchError

  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If

  BK_ribbonUI.Invalidate

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  'Call Library.showNotice(400, Err.Description, True)
End Function





'**************************************************************************************************
' * ���{�����j���[[�I�v�V����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function OptionShow(control As IRibbonControl)
  Ctl_Option.showOption
End Function

'--------------------------------------------------------------------------------------------------
Function OptionStyleImport(control As IRibbonControl)
  Call Ctl_Style.Import
End Function


'--------------------------------------------------------------------------------------------------
Function OptionStyleExport(control As IRibbonControl)
  Call Ctl_Style.Export
End Function

'--------------------------------------------------------------------------------------------------
Function OptionHighLight(control As IRibbonControl)
  Ctl_Option.HighLight
End Function

'--------------------------------------------------------------------------------------------------
Function OptionComment(control As IRibbonControl)
  Ctl_Option.Comment
End Function


'--------------------------------------------------------------------------------------------------
Function OptionShowHelp(control As IRibbonControl)
  Ctl_Option.showHelp
End Function


'--------------------------------------------------------------------------------------------------
Function OptionShowVersion(control As IRibbonControl)
  Ctl_Option.showVersion
End Function

'--------------------------------------------------------------------------------------------------
Function initialization(control As IRibbonControl)
  Ctl_Option.initialization
End Function




'**************************************************************************************************
' * ���{�����j���[[���C�ɓ���]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'--------------------------------------------------------------------------------------------------
Function FavoriteAdd(control As IRibbonControl)
  Call Ctl_Favorite.add
  
End Function

'--------------------------------------------------------------------------------------------------
Function FavoriteDetail(control As IRibbonControl)
  Call Ctl_Favorite.detail
End Function



'**************************************************************************************************
' * �J�X�^�}�C�Y
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function defaultView(control As IRibbonControl)
  Call Main.�W�����
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect(control As IRibbonControl)
  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function

'--------------------------------------------------------------------------------------------------
Function defaultViewAndSave(control As IRibbonControl)
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  ActiveWorkbook.Save
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect2(control As IRibbonControl)
  Call Main.A1�Z���I��
End Function

'--------------------------------------------------------------------------------------------------
Function dspDefaultViewSelect2AndSave(control As IRibbonControl)
  Call Main.A1�Z���I��
  ActiveWorkbook.Save
End Function





'**************************************************************************************************
' * ���{�����j���[[�u�b�N�Ǘ�]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function resetStyle(control As IRibbonControl)
  Call Ctl_Style.�X�^�C��������
End Function

'--------------------------------------------------------------------------------------------------
Function delStyle(control As IRibbonControl)
  Call Ctl_Style.�X�^�C���폜
End Function

'--------------------------------------------------------------------------------------------------
Function setStyle(control As IRibbonControl)
  Call Ctl_Style.�X�^�C���ݒ�
End Function



'--------------------------------------------------------------------------------------------------
Function ���O��`�폜(control As IRibbonControl)
  Call Ctl_Book.���O��`�폜
End Function

'--------------------------------------------------------------------------------------------------
Function getSheetList(control As IRibbonControl)
  Call Ctl_Book.�V�[�g���X�g�擾
End Function

'--------------------------------------------------------------------------------------------------
Function ���ׂĕ\��(control As IRibbonControl)
  Call Main.���ׂĕ\��
End Function


'--------------------------------------------------------------------------------------------------
Function disp_SVGA12(control As IRibbonControl)
  Call Ctl_Window.��ʃT�C�Y�ύX(612, 432)
End Function


'--------------------------------------------------------------------------------------------------
Function disp_FHD15_6(control As IRibbonControl)
  Call Ctl_Window.��ʃT�C�Y�ύX(1920, 1080)
End Function


'--------------------------------------------------------------------------------------------------
Function disp_HD15_6(control As IRibbonControl)
  Call Ctl_Window.��ʃT�C�Y�ύX(1366, 764)
End Function



'--------------------------------------------------------------------------------------------------
Function HighLight(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  Call init.setting
  
  BKh_rbPressed = pressed
  Call Library.setRegistry("Main", "HighLightFlg", pressed)
  
  Call Ctl_HighLight.showStart(ActiveCell)
  If pressed = False Then
    Call Library.delRegistry("Main", "HighLightFlg")
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function dispR1C1(control As IRibbonControl)
  Call Ctl_Sheet.R1C1�\�L
End Function


'--------------------------------------------------------------------------------------------------
Function AdjustWidth(control As IRibbonControl)
  Call Ctl_Sheet.�Z��������
End Function

'--------------------------------------------------------------------------------------------------
Function AdjustHeight(control As IRibbonControl)
  Call Ctl_Sheet.�Z����������
End Function

'--------------------------------------------------------------------------------------------------
Function AdjustHeightAndWidth(control As IRibbonControl)
  Call Ctl_Sheet.�Z��������
  Call Ctl_Sheet.�Z����������
End Function


'--------------------------------------------------------------------------------------------------
Function getAdjustWidth(control As IRibbonControl)
  Call Library.getColumnWidth
End Function



'--------------------------------------------------------------------------------------------------
Function Zoom(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKz_rbPressed = pressed
  
  
  Call init.setting
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  If pressed = False Then
    Call Application.OnKey("{F2}")
    Call Library.delRegistry("Main", "ZoomFlg")
  Else
    Call Application.OnKey("{F2}", "Ctl_Zoom.ZoomIn")
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function Zoom01(control As IRibbonControl)
  Call Ctl_Zoom.Zoom01
  
End Function




'--------------------------------------------------------------------------------------------------
Function stamp01(control As IRibbonControl)
  Call Ctl_Stamp.����_�ψ�
End Function

'--------------------------------------------------------------------------------------------------
Function stamp02(control As IRibbonControl)
  Call Ctl_Stamp.����_�m�F��
End Function

'--------------------------------------------------------------------------------------------------
Function stamp03(control As IRibbonControl)
  Call Ctl_Stamp.����_�ψ�
End Function



'--------------------------------------------------------------------------------------------------
Function confirmFormula(control As IRibbonControl, pressed As Boolean)
  Call Library.endScript
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKcf_rbPressed = pressed
  
  
  Call init.setting
  'Call Library.setRegistry("ZoomIn", ActiveWorkbook.Name, pressed)
  
  Call Ctl_Formula.�����m�F
End Function

'--------------------------------------------------------------------------------------------------
Function formatComment(control As IRibbonControl)
  Call Ctl_format.�R�����g���`
End Function

'--------------------------------------------------------------------------------------------------
Function formatMoveAndSize(control As IRibbonControl)
  Call Ctl_format.�ړ���T�C�Y�ύX������
End Function
'--------------------------------------------------------------------------------------------------
Function formatMove(control As IRibbonControl)
  Call Ctl_format.�ړ�����
End Function
'--------------------------------------------------------------------------------------------------
Function formatFreeFloating(control As IRibbonControl)
  Call Ctl_format.�ړ���T�C�Y�ύX�����Ȃ�
End Function
'--------------------------------------------------------------------------------------------------
Function MarginZero(control As IRibbonControl)
  Call Ctl_format.�]���[��
End Function







'--------------------------------------------------------------------------------------------------
Function formula01(control As IRibbonControl)
  Call Ctl_Formula.formula01
End Function

'--------------------------------------------------------------------------------------------------
Function formula02(control As IRibbonControl)
  Call Main.xxxxxxxxxx
End Function

'--------------------------------------------------------------------------------------------------
Function formula03(control As IRibbonControl)
  Call Main.xxxxxxxxxx
End Function

'--------------------------------------------------------------------------------------------------
Function saveSelectArea2Image(control As IRibbonControl)
  Call Ctl_Image.saveSelectArea2Image
End Function


'**************************************************************************************************
' * ���{�����j���[[�Z���ҏW]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function FuncCtl_Cells(control As IRibbonControl)
  Select Case control.ID
    Case "Trim"
        Call Ctl_Cells.Trim01
    Case "�����_�t�^"
      Call Ctl_Cells.�����_�t�^
    Case "�A�Ԓǉ�"
      Call Ctl_Cells.�A�Ԓǉ�
    Case "�S���p�ϊ�"
      Call Ctl_Cells.�p�����S���p�ϊ�
    Case "��������"
      Call Ctl_Cells.���������ݒ�
    Case "�R�����g�}��"
      Call Ctl_Cells.�R�����g�}��
    Case "�R�����g�폜"
      Call Ctl_Cells.�R�����g�폜
  End Select
End Function

'**************************************************************************************************
' * ���{�����j���[[�r��]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function �r��_�N���A(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�N���A
End Function
'--------------------------------------------------------------------------------------------------
Function �r��_�N���A_������_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�������폜_��
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�N���A_������_�c(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�������폜_�c
End Function


'**************************************************************************************************
' * ���{�����j���[[�r��_�\]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function �r��_�\_�j��A(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�\
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�\_�j��B(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_�i�q
  Call Library.�r��_����_����
  Call Library.�r��_����_�͂�
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�\_����(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_�i�q
End Function


'**************************************************************************************************
' * ���{�����j���[[�r��_�j��]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function �r��_�j��_����(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_����
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_����(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_����
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_��
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_�E(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_�E
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_���E(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_���E
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_��
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_��
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_�㉺(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_�㉺
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_�͂�(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_�͂�
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_�j��_�i�q(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_�j��_�i�q
End Function


'**************************************************************************************************
' * ���{�����j���[[�r��_����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function �r��_����_����(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_����
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_����_����(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_����
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_����_���E(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_���E
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_����_�㉺(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_�㉺
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_����_�͂�(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_�͂�
End Function

'--------------------------------------------------------------------------------------------------
Function �r��_����_�i�q(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_����_�i�q
End Function




'**************************************************************************************************
' * ���{�����j���[[�r��_��d��]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �r��_��d��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_��
End Function
'==================================================================================================
Function �r��_��d��_���E(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_���E
End Function

'==================================================================================================
Function �r��_��d��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_��
End Function

'==================================================================================================
Function �r��_��d��_��(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_��
End Function

'==================================================================================================
Function �r��_��d��_�㉺(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_�㉺
End Function
'==================================================================================================
Function �r��_��d��_�͂�(control As IRibbonControl)
  Call init.setting
  Call Library.�r��_��d��_�͂�
End Function





'**************************************************************************************************
' * ���{�����j���[[�T���v���f�[�^����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function makeSampleData_SelectPattern(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.�p�^�[���I��
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_DigitsInt(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���l_�����Œ�(Selection.count)
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_RangeInt(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���l_�͈�
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_FamilyName(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���O_��(Selection.count)
End Function
'--------------------------------------------------------------------------------------------------
Function makeSampleData_Name(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���O_��(Selection.count)
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_FullName(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���O_�t���l�[��(Selection.count)
End Function


'--------------------------------------------------------------------------------------------------
Function makeSampleData_Date(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���t_��(Selection.count)
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_Time(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.���t_����(Selection.count)
End Function

'--------------------------------------------------------------------------------------------------
Function makeSampleData_Datetime(control As IRibbonControl)
  Call init.setting
  Call Ctl_sampleData.����(Selection.count)
End Function



'**************************************************************************************************
' * ���{�����j���[[���̑�]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function ShrinkToFit(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.ShrinkToFit = True
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function ShrinkToUnfit(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.ShrinkToFit = False
  End If
End Function


'**************************************************************************************************
' * ���{�����j���[[RelaxTools]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function RelaxTools_get(control As IRibbonControl)
  CreateObject("WScript.Shell").run ("chrome.exe -url " & "https://software.opensquare.net/relaxtools/")
End Function


'==================================================================================================
Function actRelaxSheetManager(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!execSheetManager"
End Function

'==================================================================================================
Function RelaxTools01(control As IRibbonControl)
  Call init.setting
End Function

'==================================================================================================
'�T�C�Y���킹
Function RelaxShapes01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeSize"
End Function

'==================================================================================================
'��ʒu���킹
Function RelaxShapes02(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeTop"
End Function

'==================================================================================================
'���ʒu���킹
Function RelaxShapes03(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!sameShapeLeft"
End Function


'==================================================================================================
'�t�k�r��
Function RelaxApps01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!execSelectionFormatCheckList"
End Function








