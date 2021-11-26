Attribute VB_Name = "Ctl_Ribbon"
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
  Const funcName As String = "Ctl_Ribbon.onLoad"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  Set BK_ribbonUI = ribbon
  
  BKh_rbPressed = Library.getRegistry("Main", "HighLightFlg")
  BKz_rbPressed = Library.getRegistry("Main", "ZoomFlg")
  BKT_rbPressed = Library.getRegistry("Main", "CustomRibbon")
  
  Call Library.showDebugForm("BKh_rbPressed", CStr(BKh_rbPressed), "debug")
  Call Library.showDebugForm("BKz_rbPressed", CStr(BKz_rbPressed), "debug")
  Call Library.showDebugForm("BKT_rbPressed", CStr(BKT_rbPressed), "debug")
  
  Call Library.setRegistry("Main", "BK_ribbonUI", CStr(ObjPtr(BK_ribbonUI)))
  
  'BK_ribbonUI.ActivateTab ("Ladex")
  BK_ribbonUI.Invalidate
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

''==================================================================================================
'Function dMenuRefresh(control As IRibbonControl)
'  If BK_ribbonUI Is Nothing Then
'    #If VBA7 And Win64 Then
'      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
'    #Else
'      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
'    #End If
'  End If
'  BK_ribbonUI.Invalidate
'End Function

'==================================================================================================
Function getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  returnedVal = Library.getRegistry("Main", "CustomRibbon")
  Call RefreshRibbon
End Function

'==================================================================================================
Function RefreshRibbon()
  Const funcName As String = "Ctl_Ribbon.RefreshRibbon"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  'Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  BK_ribbonUI.Invalidate

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description & ">", "Error")
  Call Library.errorHandle
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

'**************************************************************************************************
' * �g�O���{�^���Ƀ`�F�b�N��ݒ肷��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�n�C���C�g
'==================================================================================================
Function HighLight(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.HighLight"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKh_rbPressed = pressed
  Call Library.setRegistry("Main", "HighLightFlg", pressed)
  
  Call Ctl_HighLight.showStart(ActiveCell)
  If pressed = False Then
    Call Library.delRegistry("Main", "HighLightFlg")
  End If
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function HighLightPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKh_rbPressed
End Function

' �Y�[��
'==================================================================================================
Function Zoom(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.Zoom"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  
  BKz_rbPressed = pressed
  If pressed = False Then
    Call Application.OnKey("{F2}")
    Call Library.delRegistry("Main", "ZoomFlg")
  Else
    Call Application.OnKey("{F2}", "Ctl_Zoom.ZoomIn")
  End If
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function ZoomInPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKz_rbPressed
End Function

'==================================================================================================
' �v�Z���m�F
Function confirmFormula(control As IRibbonControl, pressed As Boolean)
  Const funcName As String = "Ctl_Ribbon.confirmFormula"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  Set ctlEvent = New clsEvent
  Set ctlEvent.ExcelApplication = Application
  ctlEvent.InitializeBookSheets
  BKcf_rbPressed = pressed
  
  Call Ctl_Formula.�����m�F
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function confFormulaPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = BKcf_rbPressed
End Function


'==================================================================================================
'���C�ɓ���t�@�C�����J��
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

'**************************************************************************************************
' * ���{�����j���[�\��/��\���؂�ւ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
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

'**************************************************************************************************
' * �_�C�i�~�b�N���j���[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�V�[�g�ꗗ���j���[
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa, sheetNameID
  
'  On Error GoTo catchError
  If Workbooks.count = 0 Then
    Call MsgBox("�u�b�N���J����Ă��܂���", vbCritical, thisAppName)
    Call Library.endScript
    End
  End If
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
  'Debug.Print DOMDoc.XML
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
' ���C�ɓ��胁�j���[
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa
  
  Const funcName As String = "Ctl_Ribbon.FavoriteMenu"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
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
  'Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
'RelaxTools
Function getRelaxTools(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa

  Const funcName As String = "Ctl_Ribbon.getRelaxTools"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
    
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
    'RelaxTools�擾------------------------------
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
    
    'RelaxTools----------------------------------
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
    
    'RelaxShapes---------------------------------
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
    
    'RelaxApps-----------------------------------
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
    'RelaxTools�擾------------------------------
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
  'Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing

  BK_ribbonUI.InvalidateControl control.ID

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function selectActiveSheet(control As IRibbonControl)
  Dim sheetNameID As Integer
  Dim sheetCount As Integer
  Dim sheetName As Worksheet
  Const funcName As String = "Ctl_Ribbon.selectActiveSheet"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
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
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
'  If BK_ribbonUI Is Nothing Then
'  Else
'    BK_ribbonUI.Invalidate
'  End If
  
  Call Library.endScript
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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

'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function
'**************************************************************************************************
' * ���{�����j���[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '�����J�n--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "start")
  Call Library.startScript
  '----------------------------------------------
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  
  Select Case control.ID
    Case "Option"
      Ctl_Option.showOption
      
    Case "Favorite_detail"
      Call Ctl_Favorite.detail
    
    Case "Notation_R1C1"
      Call Ctl_Sheet.R1C1�\�L
    
    'Option--------------------------------------
    Case "Version"
      Call Ctl_Option.showVersion
    Case "Help"
      Call Ctl_Option.showHelp
    Case "OptionSheetImport"
      Call Ctl_RbnMaint.OptionSheetImport
    Case "OptionSheetExport"
      Call Ctl_RbnMaint.OptionSheetExport
    
    '�u�b�N�Ǘ�----------------------------------
    Case "resetStyle"
      Call Ctl_Style.�X�^�C��������
    Case "delStyle"
      Call Ctl_Style.�X�^�C���폜
    Case "setStyle"
      Call Ctl_Style.�X�^�C���ݒ�
    Case "del_CellNames"
      Call Ctl_Book.���O��`�폜
    Case "disp_SVGA12"
      Call Ctl_Window.��ʃT�C�Y�ύX(612, 432)
    Case "disp_HD15_6"
      Call Ctl_Window.��ʃT�C�Y�ύX(1920, 1080)
    Case "�V�[�g�ꗗ�擾"
      Call Ctl_Book.�V�[�g���X�g�擾
    
    '�V�[�g�Ǘ�----------------------------------
    Case "�Z���I��"
      Application.Goto Reference:=Range("A1"), Scroll:=True
    Case "�Z���I��_�ۑ�"
      Application.Goto Reference:=Range("A1"), Scroll:=True
      ActiveWorkbook.Save
    Case "�S�Z���\��"
      Call Ctl_Sheet.���ׂĕ\��
    Case "�Z���ƃV�[�g�I��"
      Call Ctl_Sheet.A1�Z���I��
    Case "�Z���ƃV�[�g_�ۑ�"
      Call Ctl_Sheet.A1�Z���I��
      ActiveWorkbook.Save
    Case "�W�����"
      Call Ctl_Sheet.�W�����
    
    '�Y�[��--------------------------------------
    Case "Zoom01"
      Call Ctl_Zoom.Zoom01
    
    '�Z������------------------------------------
    Case "�Z������_��"
      Call Ctl_Sheet.�Z��������
    Case "�Z������_����"
      Call Ctl_Sheet.�Z����������
    Case "�Z������_����"
      Call Ctl_Sheet.�Z��������
      Call Ctl_Sheet.�Z����������
    Case "�Z�����擾"
      Call Library.getColumnWidth
    
    '�Z���ҏW------------------------------------
    Case "Trim01"
        Call Ctl_Cells.Trim01
    Case "Trim02"
        Call Ctl_Cells.�S�󔒍폜
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
    Case "�R�����g���`"
      Call Ctl_format.�R�����g���`
    
    '�����ҏW------------------------------------
    Case "formula01"
      Call Ctl_Formula.formula01
    
    '���`------------------------------------
    Case "���`_1"
      Call Ctl_format.�ړ���T�C�Y�ύX������
    Case "���`_2"
      Call Ctl_format.�ړ�����
    Case "���`_3"
      Call Ctl_format.�ړ���T�C�Y�ύX�����Ȃ�
    Case "�]���[��"
      Call Ctl_format.�]���[��
    
    '�摜�ۑ�------------------------------------
    Case "saveImage"
      Call Ctl_Image.saveSelectArea2Image
    
    '�r��[�N���A]--------------------------------
    Case "�r��_�N���A"
      Call Library.�r��_�N���A
    Case "�r��_�N���A_������_��"
      Call Library.�r��_�������폜_��
    Case "�r��_�N���A_������_�c"
      Call Library.�r��_�������폜_�c
    
    '�r��[�\]------------------------------------
    Case "�r��_�\_����"
      Call Library.�r��_����_�i�q
    Case "�r��_�\_�j��B"
      Call Library.�r��_�\
    Case "�r��_�\_�j��C"
      Call Library.�r��_�j��_�i�q
      Call Library.�r��_����_����
      Call Library.�r��_����_�͂�
    
    '�r��[�j��]----------------------------------
    Case "�r��_�j��_����"
      Call Library.�r��_�j��_����
    Case "�r��_�j��_����"
      Call Library.�r��_�j��_����
    Case "�r��_�j��_��"
      Call Library.�r��_�j��_��
    Case "�r��_�j��_�E"
      Call Library.�r��_�j��_�E
    Case "�r��_�j��_���E"
      Call Library.�r��_�j��_���E
    Case "�r��_�j��_��"
      Call Library.�r��_�j��_��
    Case "�r��_�j��_��"
      Call Library.�r��_�j��_��
    Case "�r��_�j��_�㉺"
      Call Library.�r��_�j��_�㉺
    Case "�r��_�j��_�͂�"
      Call Library.�r��_�j��_�͂�
    Case "�r��_�j��_�i�q"
      Call Library.�r��_�j��_�i�q
    
    '�r��[����]----------------------------------
    Case "�r��_����_����"
      Call Library.�r��_����_����
    Case "�r��_����_����"
      Call Library.�r��_����_����
    Case "�r��_����_���E"
      Call Library.�r��_����_���E
    Case "�r��_����_�㉺"
      Call Library.�r��_����_�㉺
    Case "�r��_����_�͂�"
      Call Library.�r��_����_�͂�
    Case "�r��_����_�i�q"
      Call Library.�r��_����_�i�q
    
    '�r��[��d��]----------------------------------
    Case "�r��_��d��_��"
      Call Library.�r��_��d��_��
    Case "�r��_��d��_���E"
      Call Library.�r��_��d��_���E
    Case "�r��_��d��_��"
      Call Library.�r��_��d��_��
    Case "�r��_��d��_��"
      Call Library.�r��_��d��_��
    Case "�r��_��d��_�㉺"
      Call Library.�r��_��d��_�㉺
    Case "�r��_��d��_�͂�"
      Call Library.�r��_��d��_�͂�
      
    '�f�[�^����-----------------------------------
    Case "�A�Ԑݒ�"
      Call Ctl_Cells.�A�Ԑݒ�
    Case "�A�Ԑ���"
      Call Ctl_Cells.�A�Ԓǉ�
    Case "�����Œ萔�l"
      Call Ctl_sampleData.���l_�����Œ�(Selection.count)
    Case "�͈͎w�萔�l"
      Call Ctl_sampleData.���l_�͈�
    Case "��"
      Call Ctl_sampleData.���O_��(Selection.count)
    Case "��"
      Call Ctl_sampleData.���O_��(Selection.count)
    Case "����"
      Call Ctl_sampleData.���O_�t���l�[��(Selection.count)
    Case "���t"
      Call Ctl_sampleData.���t_��(Selection.count)
    Case "����"
      Call Ctl_sampleData.���t_����(Selection.count)
    Case "����"
      Call Ctl_sampleData.����(Selection.count)
    
    
    Case Else
      Call Library.showDebugForm("���{�����j���[�Ȃ�", control.ID, "Error")
  End Select
  
  '�����I��--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm("  ", , "end")
  Call init.unsetting
  '----------------------------------------------
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
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


