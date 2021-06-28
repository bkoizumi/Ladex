Attribute VB_Name = "init"

'���[�N�u�b�N�p�ϐ�------------------------------
Public BK_ThisBook   As Workbook
Public targetBook As Workbook


'���[�N�V�[�g�p�ϐ�------------------------------
Public BK_sheetSetting    As Worksheet
Public BK_sheetNotice     As Worksheet
Public BK_sheetStyle      As Worksheet
Public BK_sheetTestData   As Worksheet
Public BK_sheetRibbon     As Worksheet
Public BK_sheetFavorite   As Worksheet
Public BK_sheetStamp      As Worksheet
Public BK_sheetHighLight  As Worksheet


'�O���[�o���ϐ�----------------------------------
Public Const thisAppName = "Ladex"
Public Const thisAppVersion = "0.0.4.0"
Public FuncName           As String
Public resetVal           As String

'���W�X�g���o�^�p�T�u�L�[
Public Const RegistryKey  As String = "Ladex"
Public RegistrySubKey     As String


'�ݒ�l�ێ�
Public BK_setVal          As Object
Public sampleDataList     As Object


'�t�@�C��/�f�B���N�g���֘A
Public logFile            As String
Public LadexDir           As String
Public LadexPreDir        As String


'�������Ԍv���p
Public StartTime          As Date
Public StopTime           As Date



'���{���֘A--------------------------------------
Public BK_ribbonUI        As Office.IRibbonUI
Public BK_ribbonVal       As Object
Public BKT_rbPressed      As Boolean

Public BKh_rbPressed      As Boolean
Public BKz_rbPressed      As Boolean
Public BKcf_rbPressed     As Boolean



'���[�U�[�֐��֘A--------------------------------
Public arryHollyday()     As Date

'�Y�[���֘A--------------------------------------
Public defaultZoomInVal   As String


'**************************************************************************************************
' * �ݒ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function usetting()

  Set BK_ThisBook = Nothing
  
  '���[�N�V�[�g���̐ݒ�
  Set BK_sheetSetting = Nothing
  Set BK_sheetNotice = Nothing
  Set BK_sheetStyle = Nothing
  Set BK_sheetTestData = Nothing
  Set BK_sheetRibbon = Nothing
  Set BK_sheetFavorite = Nothing

  '�ݒ�l�ǂݍ���
  Set BK_setVal = Nothing
  Set BK_ribbonVal = Nothing
  
  logFile = ""
End Function


'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  
  On Error GoTo catchError
'  ThisWorkbook.Save

  '���W�X�g���֘A�ݒ�------------------------------------------------------------------------------
  RegistrySubKey = "Main"
  
  If LadexDir = "" Or reCheckFlg = True Then
    Call usetting
  Else
    Exit Function
  End If

  '�u�b�N�̐ݒ�
  Set BK_ThisBook = ThisWorkbook
  
  '���[�N�V�[�g���̐ݒ�
  Set BK_sheetSetting = BK_ThisBook.Worksheets("�ݒ�")
  Set BK_sheetNotice = BK_ThisBook.Worksheets("Notice")
  Set BK_sheetStyle = BK_ThisBook.Worksheets("Style")
  Set BK_sheetTestData = BK_ThisBook.Worksheets("testData")
'  Set BK_sheetRibbon = BK_ThisBook.Worksheets("Ribbon")
  Set BK_sheetFavorite = BK_ThisBook.Worksheets("Favorite")
  Set BK_sheetStamp = BK_ThisBook.Worksheets("Stamp")
  Set BK_sheetHighLight = BK_ThisBook.Worksheets("HighLight")
 
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
        
  '�ݒ�l�ǂݍ���----------------------------------------------------------------------------------
  Set BK_setVal = Nothing
  Set BK_setVal = CreateObject("Scripting.Dictionary")
  BK_setVal.add "debugMode", "develop"
  
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetSetting.Range("A" & line) <> "" Then
      BK_setVal.add BK_sheetSetting.Range("A" & line).Text, BK_sheetSetting.Range("B" & line).Text
    End If
  Next
  
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")

  LadexDir = wsh.SpecialFolders("AppData") & "\Ladex"
  LadexPreDir = wsh.SpecialFolders("AppData") & "\Ladex\Preview"
  
  
'  Set BK_ribbonVal = Nothing
'  Set BK_ribbonVal = CreateObject("Scripting.Dictionary")
'  For line = 2 To BK_sheetRibbon.Cells(Rows.count, 1).End(xlUp).Row
'    If BK_sheetRibbon.Range("A" & line) <> "" Then
'      BK_ribbonVal.add "Lbl_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("B" & line).Text
'      BK_ribbonVal.add "Act_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("C" & line).Text
'      BK_ribbonVal.add "Sup_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("D" & line).Text
'      BK_ribbonVal.add "Dec_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("E" & line).Text
'      BK_ribbonVal.add "Siz_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("F" & line).Text
'      BK_ribbonVal.add "Img_" & BK_sheetRibbon.Range("A" & line).Text, BK_sheetRibbon.Range("G" & line).Text
'    End If
'  Next
  
  
  Exit Function
  
'�G���[������=====================================================================================
catchError:
  
End Function


'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError

  '���O�̒�`���폜
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" And _
      Not Name.Name Like "Slc*" And Not Name.Name Like "Pvt*" And Not Name.Name Like "Tbl*" Then
      Name.delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To BK_sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If BK_sheetSetting.Range("A" & line) <> "" Then
      BK_sheetSetting.Range("B" & line).Name = BK_sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book�p�̐ݒ�
  BK_sheetSetting.Range("D3:D" & BK_sheetSetting.Cells(Rows.count, 6).End(xlUp).Row).Name = BK_sheetSetting.Range("D2")
  

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

