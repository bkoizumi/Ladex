Attribute VB_Name = "Ctl_Ribbon"
Option Explicit

Private Ctl_Event As New Ctl_Event

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
  Call Library.showDebugForm(funcName, , "start")
  '----------------------------------------------
  
  Set BK_ribbonUI = ribbon
  
  BKh_rbPressed = Library.getRegistry("Main", "HighLightFlg", "Boolean")
  BKz_rbPressed = Library.getRegistry("Main", "ZoomFlg", "Boolean")
  BKT_rbPressed = Library.getRegistry("Main", "CustomRibbon", "Boolean")
  
  Call Library.showDebugForm("BKh_rbPressed", CStr(BKh_rbPressed), "debug")
  Call Library.showDebugForm("BKz_rbPressed", CStr(BKz_rbPressed), "debug")
  Call Library.showDebugForm("BKT_rbPressed", CStr(BKT_rbPressed), "debug")
  
  Call Library.setRegistry("Main", "BK_ribbonUI", CStr(ObjPtr(BK_ribbonUI)))
  Call Main.InitializeBook
  
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
  'Call Library.showDebugForm(funcName, , "start1")
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
  Dim targetBook  As Workbook
  Dim targetSheet As Worksheet
  
  Const funcName As String = "Ctl_Ribbon.HighLight"
  
  '�����J�n--------------------------------------
  runFlg = True
  '  On Error GoTo catchError
  Call init.setting
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
  BKh_rbPressed = pressed
  
  If pressed = False Then
      If Library.chkShapeName("HighLight_X", ActiveSheet) = True Then
        ActiveSheet.Shapes("HighLight_X").delete
      End If
      If Library.chkShapeName("HighLight_Y", ActiveSheet) = True Then
        ActiveSheet.Shapes("HighLight_Y").delete
      End If
    
    Call Library.setRegistry("Main", "HighLightFlg", pressed)
    Call Library.delRegistry("targetInfo", "HighLight_Book")
    Call Library.delRegistry("targetInfo", "HighLight_Sheet")
  
  Else
    Call Library.setRegistry("Main", "HighLightFlg", pressed)
    Call Library.setRegistry("targetInfo", "HighLight_Book", ActiveWorkbook.Name)
    Call Library.setRegistry("targetInfo", "HighLight_Sheet", ActiveSheet.Name)
    
    Call Ctl_HighLight.showStart(ActiveCell)
  End If
  Call Library.endScript
  
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
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  Call Library.setRegistry("Main", "ZoomFlg", pressed)
  
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  
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
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
  '----------------------------------------------
  Set Ctl_Event = New Ctl_Event
  Set Ctl_Event.ExcelApplication = Application
  Ctl_Event.InitializeBookSheets
  BKcf_rbPressed = pressed
  
  
  If BKcf_rbPressed = True Then
    Call Library.setRegistry("targetInfo", "Formula_Book", ActiveWorkbook.Name)
    Call Library.setRegistry("targetInfo", "Formula_Sheet", ActiveSheet.Name)
  Else
    Call Library.delRegistry("targetInfo", "Formula_Book")
    Call Library.delRegistry("targetInfo", "Formula_Sheet")
  End If
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
Function FavoriteFileOpen(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  Dim objFso As New FileSystemObject
  Const funcName As String = "Ctl_Ribbon.FavoriteFileOpen"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  runFlg = True
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------
  
  fileNamePath = Library.getRegistry("FavoriteList", Replace(control.ID, ".-.", "<L|>"))
  
  If Library.chkFileExists(fileNamePath) Then
    If Library.chkBookOpened(fileNamePath) = True Then
      Call Library.showNotice(415, "", True)
    Else
      Select Case objFso.GetExtensionName(fileNamePath)
        Case "xls", "xlsx", "xlsm", "xlam"
          Workbooks.Open fileName:=fileNamePath
        Case Else
          CreateObject("Shell.Application").ShellExecute fileNamePath
      End Select
    End If
  Else
    Call Library.showNotice(404, fileNamePath, True)
  End If
  
  '�����I��--------------------------------------
  Call Library.endScript
  Call Library.showDebugForm(funcName, , "end")
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'���C�ɓ���t�@�C���ǉ�
Function FavoriteAddFile(control As IRibbonControl, ByRef returnedVal)
  Dim fileNamePath As String
  Dim line As Long
  Dim setCategory As Long

  Const funcName As String = "Ctl_Ribbon.FavoriteAddFile"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  runFlg = True
  Call Library.showDebugForm(funcName, , "start1")
  Call Library.startScript
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------

  setCategory = Replace(control.ID, "M_FavoriteCategory", "")
  Call Library.showDebugForm("setCategory", setCategory, "debug")


  Call Ctl_Favorite.�ǉ�(setCategory, ActiveWorkbook.FullName)

  Call Library.delSheetData(LadexSh_Favorite)

  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")

  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
'���C�ɓ���t�@�C���ǉ��J�e�S���[���j���[�\��
Function FavoritesToAdd(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object, CategoryMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa, tmp, Category
  Dim categoryName As String, oldCategoryName As String
  
  Const funcName As String = "Ctl_Ribbon.AddToFavorites"

  '�����J�n--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  
  Call Ctl_Favorite.chkDebugMode
  '----------------------------------------------
  
  If Library.Book�̏�Ԋm�F = False Then
    Call MsgBox("�u�b�N���J����Ă��܂���", vbCritical, thisAppName)
    Call Library.errorHandle
    End
  End If
  
  Call Ctl_Favorite.���X�g�擾
  
  If BK_ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set BK_ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "BK_ribbonUI")))
    #Else
      Set BK_ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "BK_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menu�̍쐬

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
  With MenuSepa
    .setAttribute "id", "MS_���C�ɓ���ǉ��J�e�S���["
    .setAttribute "title", "���C�ɓ���ǉ��J�e�S���["
  End With
  Menu.appendChild MenuSepa
  Set MenuSepa = Nothing
  
  
  endLine = targetSheet.Cells(Rows.count, 1).End(xlUp).Row
  If IsEmpty(tmp) Then
    targetSheet.Range("A1") = "Category01"
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "M_FavoriteCategory" & 1
      .setAttribute "label", "Category01"
      .setAttribute "imageMso", "AddFolderToFavorites"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteAddFile"
    End With

    Menu.appendChild Button
    Set Button = Nothing
  Else
    Call Library.Sort_QuickSort(tmp, LBound(tmp), UBound(tmp), 0)
    oldCategoryName = ""
    
    For i = 0 To UBound(tmp)
      categoryName = Split(tmp(i, 0), "<L|>")(0)
      Call Library.showDebugForm("categoryName", categoryName, "debug")
      
      If oldCategoryName <> categoryName Then
        Set Button = DOMDoc.createElement("button")
        With Button
          .setAttribute "id", "M_FavoriteCategory" & line
          .setAttribute "label", categoryName
          .setAttribute "imageMso", "AddFolderToFavorites"
          .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteAddFile"
        End With
    
        Menu.appendChild Button
        Set Button = Nothing
        oldCategoryName = categoryName
        line = line + 1
      End If
    Next
    
  End If
  
  DOMDoc.appendChild Menu
  returnedVal = DOMDoc.XML
  Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
  
  Set CategoryMenu = Nothing
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
  '�����I��--------------------------------------
  Call Library.showDebugForm(funcName, , "end")
  Call init.resetGlobalVal
  Exit Function
  '----------------------------------------------

'�G���[������------------------------------------
catchError:
  Set Menu = Nothing
  Set DOMDoc = Nothing
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
' ���C�ɓ��胁�j���[
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object, CategoryMenu As Object, CategorymenuSeparator As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFso As New FileSystemObject
  Dim MenuSepa, tmp, Category
  
  Const funcName As String = "Ctl_Ribbon.FavoriteMenu"

  '�����J�n--------------------------------------
  runFlg = True
'  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

'  Call Ctl_Favorite.getList
'  endLine = LadexSh_Favorite.Cells(Rows.count, 1).End(xlUp).Row
  tmp = GetAllSettings(thisAppName, "FavoriteList")
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
  With MenuSepa
    .setAttribute "id", "MS_�J�e�S���[�ꗗ"
    .setAttribute "title", "�J�e�S���[�ꗗ"
  End With
  Menu.appendChild MenuSepa
  Set MenuSepa = Nothing
  
  If Not IsEmpty(tmp) Then
    For line = 0 To UBound(tmp)
      Category = Split(tmp(line, 0), "<L|>")
      
      If Category(1) = 0 Then
        If line <> 0 Then
          Menu.appendChild CategoryMenu
        End If
        
        Set CategoryMenu = DOMDoc.createElement("menu")
        With CategoryMenu
          .setAttribute "id", "M_FavoriteCategory" & Category(0)
          .setAttribute "label", Category(0)
          .setAttribute "imageMso", "AddFolderToFavorites"
        End With
      End If
    
      If tmp(line, 1) <> "" Then
        Set Button = DOMDoc.createElement("button")
        With Button
          .setAttribute "id", Replace(tmp(line, 0), "<L|>", ".-.")
          .setAttribute "label", objFso.getFileName(tmp(line, 1))
          
          '�A�C�R���̐ݒ�
          Select Case objFso.GetExtensionName(tmp(line, 1))
            Case "xlsm", "xlsx", "xlam"
              .setAttribute "imageMso", "MicrosoftExcel"
            Case "xls"
              .setAttribute "imageMso", "FileSaveAsExcel97_2003"
            
            Case "pdf"
              .setAttribute "imageMso", "FileSaveAsPdf"
            Case "docs"
              .setAttribute "imageMso", "FileSaveAsWordDocx"
            Case "text"
              .setAttribute "imageMso", "FileNewContext"
            Case "accdb"
              .setAttribute "imageMso", "MicrosoftAccess"
            Case "pptx"
              .setAttribute "imageMso", "MicrosoftPowerPoint"
            Case "csv"
              .setAttribute "imageMso", "FileNewContext"
            Case "html"
              .setAttribute "imageMso", "GroupWebPageNavigation"
            
            Case Else
              .setAttribute "imageMso", "FileNewContext"
              Call Library.showDebugForm("���C�ɓ���A�C�R��", objFso.GetExtensionName(tmp(line, 1)), "Error")
          End Select
          
          
          .setAttribute "supertip", tmp(line, 1)
          .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.FavoriteFileOpen"
        End With
        CategoryMenu.appendChild Button
        Set Button = Nothing
      End If
    Next
    Menu.appendChild CategoryMenu

  Else
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "���o�^"
      .setAttribute "label", "���o�^"
      .setAttribute "imageMso", "FileNewContext"
      .setAttribute "supertip", "���o�^"
    End With
    Menu.appendChild Button
  End If
  DOMDoc.appendChild Menu
  returnedVal = DOMDoc.XML
'  Call Library.showDebugForm("DOMDoc.XML", DOMDoc.XML, "debug")
    
  Set CategoryMenu = Nothing
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"

  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "�V�[�g�Ǘ�"
      .setAttribute "title", "�V�[�g�Ǘ�"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "�V�[�g�Ǘ��\��"
      .setAttribute "label", "�V�[�g�Ǘ�"
      .setAttribute "supertip", "�V�[�g�Ǘ�"
      
      .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Menu.ladex_�V�[�g�Ǘ�_�t�H�[���\��"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
      With MenuSepa
        .setAttribute "id", "M_RelaxTools"
        .setAttribute "title", "RelaxTools�𗘗p"
      End With
      Menu.appendChild MenuSepa
      Set MenuSepa = Nothing

      Set Button = DOMDoc.createElement("button")
      With Button
        .setAttribute "id", "RelaxTools"
        .setAttribute "label", "RelaxTools"
        .setAttribute "supertip", "RelaxTools�̃V�[�g�Ǘ����N��"
        
        .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
        .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
      End With
      Menu.appendChild Button
      Set Button = Nothing
  End If
  
  Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "sheetID_0"
      .setAttribute "title", ActiveWorkbook.Name
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
  
  
  
  For Each sheetName In ActiveWorkbook.Sheets
    Set Button = DOMDoc.createElement("button")
    With Button
      sheetNameID = sheetName.Name
      .setAttribute "id", "sheetID_" & sheetName.Index
      .setAttribute "label", sheetName.Name
    
      If ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
        .setAttribute "supertip", "�A�N�e�B�u�V�[�g"
        .setAttribute "imageMso", "ExcelSpreadsheetInsert"
        
      ElseIf Sheets(sheetName.Name).Visible = True Then
       '.SetAttribute "supertip", "�A�N�e�B�u�V�[�g"
        .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      
      ElseIf Sheets(sheetName.Name).Visible = 0 Then
        .setAttribute "supertip", "��\���V�[�g"
        .setAttribute "imageMso", "SheetProtect"
      
      ElseIf Sheets(sheetName.Name).Visible = 2 Then
        .setAttribute "supertip", "�}�N���ɂ���\���V�[�g"
        .setAttribute "imageMso", "ReviewProtectWorkbook"
      
      End If
      
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.selectActiveSheet"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  Next

  DOMDoc.appendChild Menu
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
'RelaxTools
Function getRelaxTools(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa

  Const funcName As String = "Ctl_Ribbon.getRelaxTools"
  
  '�����J�n--------------------------------------
  runFlg = True
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start1")
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

  Menu.setAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.setAttribute "itemSize", "normal"
  
  If Library.chkFileExists(Application.UserLibraryPath & RelaxTools) = True Then
    'RelaxTools�擾------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxToolsGet"
      .setAttribute "title", "RelaxTool�̍X�V"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools_get"
      .setAttribute "label", "RelaxTool�̍X�V"
      .setAttribute "image", "RelaxToolsLogo"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    'RelaxTools----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxTools"
      .setAttribute "title", "RelaxTools�𗘗p"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools01"
      .setAttribute "label", "�V�[�g�Ǘ�"
      .setAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools02"
      .setAttribute "label", "�������t���b�V��"
      '.SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools02"
    End With
    Menu.appendChild Button
    Set Button = Nothing

    'RelaxShapes---------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxShapes"
      .setAttribute "title", "RelaxShapes�𗘗p"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing

    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes01"
      .setAttribute "label", "�T�C�Y���킹"
      .setAttribute "imageMso", "ShapesDuplicate"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes02"
      .setAttribute "label", "��ʒu���킹"
      .setAttribute "imageMso", "ObjectsAlignTop"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes02"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxShapes03"
      .setAttribute "label", "���ʒu���킹"
      .setAttribute "imageMso", "ObjectsAlignLeft"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxShapes03"
    End With
    Menu.appendChild Button
    Set Button = Nothing
    
    'RelaxApps-----------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxApps"
      .setAttribute "title", "RelaxApps�𗘗p"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxApps01"
      .setAttribute "label", "�t�k�r��"
      .setAttribute "imageMso", "BorderDrawGrid"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxApps01"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  Else
    'RelaxTools�擾------------------------------
    Set MenuSepa = DOMDoc.createElement("menuSeparator")
    With MenuSepa
      .setAttribute "id", "M_RelaxTools"
      .setAttribute "title", "RelaxTool�����"
    End With
    Menu.appendChild MenuSepa
    Set MenuSepa = Nothing
    
    Set Button = DOMDoc.createElement("button")
    With Button
      .setAttribute "id", "RelaxTools_get"
      .setAttribute "label", "RelaxTool�����"
      .setAttribute "image", "RelaxToolsLogo"
      .setAttribute "onAction", "Ladex.xlam!Ctl_Ribbon.RelaxTools_get"
    End With
    Menu.appendChild Button
    Set Button = Nothing
  End If

  DOMDoc.appendChild Menu
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
  runFlg = True
  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm(funcName, , "start1")
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
  
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  
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

'**************************************************************************************************
' * ���{�����j���[[���̑�]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  Dim slctCells
  Dim slctCnt As Long
  Const funcName As String = "Ctl_Ribbon.setCenter"

  '�����J�n--------------------------------------
  PrgP_Max = 2
  slctCnt = 1
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  '----------------------------------------------
  For Each slctCells In Selection
    If TypeName(Selection) = "Range" Then
      Selection.HorizontalAlignment = xlCenterAcrossSelection
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, slctCnt, Selection.count, "")
      slctCnt = slctCnt + 1
    End If
  Next
  
  
  '�����I��--------------------------------------
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.resetGlobalVal
    Call Library.showDebugForm(funcName, , "end")
    
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'--------------------------------------------------------------------------------------------------
Function setShrinkToFit(control As IRibbonControl)
  Call init.setting
  If TypeName(Selection) = "Range" Then
    Selection.ShrinkToFit = True
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function unsetShrinkToFit(control As IRibbonControl)
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
Function RelaxTools01(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!execSheetManager"
End Function

'==================================================================================================
Function RelaxTools02(control As IRibbonControl)
  Call init.setting
  Application.run "'" & Application.UserLibraryPath & RelaxTools & "'!basSelection.execSelectionToFormula"
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

'**************************************************************************************************
' * ���{�����j���[[�J�X�^�}�C�Y]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function Ctl_Function(control As IRibbonControl)
  Const funcName As String = "Ctl_Ribbon.Ctl_Function"
  
  '�����J�n--------------------------------------
  runFlg = True
  PrgP_Max = 2
  PrgP_Cnt = 0

  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  Call Library.showDebugForm("control.ID", control.ID, "debug")
  '----------------------------------------------
  
  Select Case control.ID
    Case "Favorite_detail"
      Call Ctl_Favorite.�ڍו\��
      
    Case "���C�ɓ���ǉ�"
      Call Ctl_Favorite.�ǉ�
    
    'Option--------------------------------------
    Case "Option"
      Call Ctl_Option.showOption
      
    Case "�X�^�C���o��"
      Call Ctl_Style.Export
    Case "�X�^�C���捞"
      Call Ctl_Style.Import
    
    Case "OptionHighLight"
      Ctl_Option.HighLight
    
    Case "OptionComment"
      Ctl_Option.Comment
    
    Case "Version"
      Call Ctl_Option.showVersion
    
    Case "Help"
      Call Ctl_Option.showHelp
    
    Case "OptionAddin����"
      Workbooks(ThisWorkbook.Name).IsAddin = False
    
    Case "OptionAddin��"
      Workbooks(ThisWorkbook.Name).IsAddin = True
      ThisWorkbook.Save
    
    '�u�b�N�Ǘ�----------------------------------
    Case "Notation_R1C1"
      Call Ctl_Book.R1C1�\�L
    
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
    
    Case "����͈͕\��"
      Call Ctl_Book.����͈͂̓_����\��
    Case "����͈͔�\��"
      Call Ctl_Book.����͈͂̓_�����\��
    
    
    Case "�Z���ƃV�[�g�I��"
      Call Ctl_Book.A1�Z���I��
    Case "�Z���ƃV�[�g_�ۑ�"
      Call Ctl_Book.A1�Z���I��
    Case "�W�����"
      Call Ctl_Book.�W�����
    Case "�A���V�[�g�ǉ�"
      Call Ctl_Book.�A���V�[�g�ǉ�
    
    
    '�V�[�g�Ǘ�----------------------------------
    Case "�Z���I��"
      Application.GoTo Reference:=Range("A1"), Scroll:=True
    Case "�Z���I��_�ۑ�"
      Application.GoTo Reference:=Range("A1"), Scroll:=True
      ActiveWorkbook.Save
    Case "�S�Z���\��"
      Call Ctl_Sheet.���ׂĕ\��
      ActiveWorkbook.Save
    Case "�V�[�g�Ǘ�"
      Call Ctl_Sheet.�V�[�g�Ǘ�_�t�H�[���\��
    Case "���ݒ�"
      Call Ctl_Sheet.���ݒ�
    Case "�����ݒ�"
      Call Ctl_Sheet.�����ݒ�
    Case "���ƍ�������"
      Call Ctl_Sheet.���ݒ�
      Call Ctl_Sheet.�����ݒ�
    Case "�̍وꊇ�ύX"
      Call Ctl_Sheet.�̍وꊇ�ύX
    Case "�t�H���g�ꊇ�ύX"
      Call Ctl_Sheet.�w��t�H���g�ɐݒ�
    
    
    '���̑��Ǘ�--------------------------------
    Case "�t�@�C���Ǘ�_���擾"
      Call Ctl_File.�t�@�C���p�X���
    
    Case "�t�@�C���Ǘ�_�t�H���_����"
      Call Ctl_File.�t�H���_����
    
    Case "�t�@�C���Ǘ�_�摜�\�t��"
      Call Ctl_File.�摜�\�t��
    
    Case "QR�R�[�h����"
      Call Ctl_shap.QR�R�[�h����
    
    
    Case "�J�X�^��01"
      Call Ctl_�J�X�^��.�J�X�^��01
    Case "�J�X�^��02"
      Call Ctl_�J�X�^��.�J�X�^��02
    Case "�J�X�^��03"
      Call Ctl_�J�X�^��.�J�X�^��03
    Case "�J�X�^��04"
      Call Ctl_�J�X�^��.�J�X�^��04
    Case "�J�X�^��05"
      Call Ctl_�J�X�^��.�J�X�^��05
    
    Case "�͂�_�m�F��"
      Call Ctl_Stamp.�m�F��
    Case "�͂�_���O"
      Call Ctl_Stamp.���O
    Case "�͂�_��"
      Call Ctl_Stamp.�ψ�
      
    Case "�p�X���[�h����"
      Frm_mkPasswd.Show vbModeless
      
    
    '�Y�[��--------------------------------------
    Case "�S��ʕ\��"
      Call Ctl_Zoom.�S��ʕ\��
      
    Case "�w��{��"
      Call Ctl_Zoom.�w��{��
      
      
      
      
    '�Z������------------------------------------
    Case "�Z������_��"
      Call Ctl_Cells.�Z��������
    Case "�Z������_����"
      Call Ctl_Cells.�Z����������
    Case "�Z������_����"
      Call Ctl_Cells.�Z��������
      Call Ctl_Cells.�Z����������
    Case "�Z�����擾"
      Call Library.getColumnWidth
    
    Case "�����Œ�ݒ�_15"
      Call Ctl_Cells.�Z�������Œ�ݒ�(15)
      
    Case "�����Œ�ݒ�_30"
      Call Ctl_Cells.�Z�������Œ�ݒ�(30)
    
    
    
    '�Z���ҏW------------------------------------
    Case "Trim01"
        Call Ctl_Cells.�O��̃X�y�[�X�폜
    Case "Trim02"
        Call Ctl_Cells.�S�󔒍폜
    Case "�����_�t�^"
      Call Ctl_Cells.�����_�t�^
    Case "�A�Ԓǉ�"
      Call Ctl_Cells.�A�Ԓǉ�
    Case "�S���p�ϊ�"
      Call Ctl_Cells.�p�����S�˔��p�ϊ�
      
    Case "���S�p�ϊ�"
      Call Ctl_Cells.�p�������ˑS�p�ϊ�
    
    Case "����ϊ�"
      Call Ctl_Cells.���ˑ�ϕϊ�
    Case "�召�ϊ�"
      Call Ctl_Cells.��ˏ��ϕϊ�
    
    Case "�ې��l�𐔒l�ɕϊ�"
      Call Ctl_Cells.�ې����ː��l
      
    Case "���l���ې����ɕϊ�"
      Call Ctl_Cells.���l�ˊې���
      
    Case "URL�G���R�[�h"
      Call Ctl_Cells.URL�G���R�[�h
    Case "URL�f�R�[�h"
      Call Ctl_Cells.URL�f�R�[�h
      
    Case "Unicode�G�X�P�[�v"
      Call Ctl_Cells.Unicode�G�X�P�[�v
    Case "Unicode�A���G�X�P�[�v"
      Call Ctl_Cells.Unicode�A���G�X�P�[�v
      
      
    Case "delLinefeed"
      Call Ctl_Cells.���s�폜
    Case "�萔�폜"
      Call Ctl_Cells.�萔�폜
      
      
    Case "��������"
      Call Ctl_Cells.���������ݒ�
    Case "�R�����g�}��"
      Call Ctl_Cells.�R�����g�}��
    Case "�R�����g�폜"
      Call Ctl_Cells.�R�����g�폜
    Case "�R�����g���`"
      Call Ctl_format.�R�����g���`
    
    Case "�s�����ւ��ē\�t��"
      Call Ctl_Cells.�s������ւ��ē\�t��
    Case "�[������"
      Call Ctl_Cells.�[������
    
    
    '�����ҏW------------------------------------
    Case "�����ǉ�_�G���[�h�~_��"
      Call Ctl_Formula.�G���[�h�~_��
    Case "�����ǉ�_�G���[�h�~_�[��"
      Call Ctl_Formula.�G���[�h�~_�[��
      
    Case "�����ǉ�_�[����\��"
      Call Ctl_Formula.�[����\��
    
    Case "�����ǉ�_row-XX�ǉ�"
      Call Ctl_Formula.�s�ԍ��ǉ�
      
    Case "�����ǉ�_�V�[�g���ǉ�"
      Call Ctl_Formula.�����}��("SheetName")
      
    
    '���`------------------------------------
    Case "���`_1"
      Call Ctl_format.�ړ���T�C�Y�ύX������
    Case "���`_2"
      Call Ctl_format.�ړ�����
    Case "���`_3"
      Call Ctl_format.�ړ���T�C�Y�ύX�����Ȃ�
    Case "�㉺�]���[��"
      Call Ctl_format.�㉺�]���[��
    Case "���E�]���[��"
      Call Ctl_format.���E�]���[��
    Case "�����T�C�Y���҂�����"
      Call Ctl_shap.�����T�C�Y���҂�����
    Case "�Z�����̒����ɔz�u"
      Call Ctl_shap.�Z���̒����ɔz�u
    
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
    Case "�r��_��d��_�E"
      Call Library.�r��_��d��_�E
    
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
      Call Ctl_sampleData.���l_�����Œ�(Selection.Rows.count)
    Case "�͈͎w�萔�l"
      Call Ctl_sampleData.���l_�͈�
    Case "��"
      Call Ctl_sampleData.���O_��(Selection.Rows.count)
    Case "��"
      Call Ctl_sampleData.���O_��(Selection.Rows.count)
    Case "����"
      Call Ctl_sampleData.���O_�t���l�[��(Selection.Rows.count)
    Case "���t"
      Call Ctl_sampleData.���t_��(Selection.Rows.count)
    Case "����"
      Call Ctl_sampleData.���t_����(Selection.Rows.count)
    Case "����"
      Call Ctl_sampleData.����(Selection.Rows.count)
    Case "����"
      Call Ctl_sampleData.���̑�_����(25)
    Case "�p�^�[���I��"
      Call Ctl_sampleData.�p�^�[���I��(Selection.Rows.count)
    

    Case Else
      Call Library.showDebugForm("���{�����j���[�Ȃ�", control.ID, "Error")
      Call Library.showNotice(406, "���{�����j���[�Ȃ��F" & control.ID, True)
  End Select
  
  '�����I��--------------------------------------
  Call Library.endScript
  Call init.resetGlobalVal
  Call Library.showDebugForm(funcName, , "end")
  '----------------------------------------------
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

