Attribute VB_Name = "init"
Option Explicit


'���[�N�u�b�N�p�ϐ�------------------------------
Public BK_ThisBook          As Workbook
Public targetBook           As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
Public targetSheet          As Worksheet

'Public LadexSh_Config      As Worksheet
'Public LadexSh_Notice       As Worksheet
'Public LadexSh_Style        As Worksheet
'Public LadexSh_TestData     As Worksheet
'Public LadexSh_Ribbon       As Worksheet
'Public LadexSh_Favorite     As Worksheet
'Public LadexSh_Stamp        As Worksheet
'Public LadexSh_HighLight    As Worksheet
'Public LadexSh_Help         As Worksheet
'Public LadexSh_Function     As Worksheet
'Public LadexSh_SheetList    As Worksheet
'Public LadexSh_InputData    As Worksheet

'�O���[�o���ϐ�----------------------------------
Public Const thisAppName    As String = "Ladex"
Public Const thisAppVersion As String = "1.3.1.0"
Public Const RelaxTools     As String = "Relaxtools.xlam"

Public funcName             As String
Public resetVal             As String
Public runFlg               As Boolean
Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
'Public LogLevel             As Long



'���W�X�g���o�^�p�L�[----------------------------
Public Const RegistryKey    As String = "Ladex"
Public RegistrySubKey       As String


'�ݒ�l�ێ�--------------------------------------
Public BK_setVal            As Object
Public sampleDataList       As Object


'�t�@�C��/�f�B���N�g���֘A-----------------------
Public logFile              As String
Public LadexDir             As String


'�������Ԍv���p----------------------------------
Public StartTime            As Date
Public StopTime             As Date



'���{���֘A--------------------------------------
Public BK_ribbonUI          As Office.IRibbonUI
Public BK_ribbonVal         As Object
Public BKT_rbPressed        As Boolean

Public BKh_rbPressed        As Boolean
Public BKz_rbPressed        As Boolean
Public BKcf_rbPressed       As Boolean



'���[�U�[�֐��֘A--------------------------------
Public arryHollyday()       As Date

'�Y�[���֘A--------------------------------------
Public defaultZoomInVal     As String




'**************************************************************************************************
' * �ݒ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "init.unsetting"

  Set BK_ThisBook = Nothing
  
  '���[�N�V�[�g���̐ݒ�
'  Set LadexSh_Config = Nothing
'  Set LadexSh_Notice = Nothing
'  Set LadexSh_Style = Nothing
'  Set LadexSh_TestData = Nothing
'  Set LadexSh_Ribbon = Nothing
'  Set LadexSh_Favorite = Nothing

  '�ݒ�l�ǂݍ���
  Set BK_setVal = Nothing
  Set BK_ribbonVal = Nothing
  
  logFile = ""
  LadexDir = ""
  
  If flg = True Then
    runFlg = False
  End If
  
  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
  Const funcName As String = "init.setting"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
'  ThisWorkbook.Save
'  If Workbooks.count = 0 Then
'    Call MsgBox("�u�b�N���J����Ă��܂���", vbCritical, thisAppName)
'    Call Library.endScript
'    End
'  End If
  '----------------------------------------------

  If LadexDir = "" Or BK_setVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  '���W�X�g���֘A
  RegistrySubKey = "Main"
  
  '�u�b�N�̐ݒ�
  Set BK_ThisBook = ThisWorkbook
  
  '���[�N�V�[�g���̐ݒ�
'  Set LadexSh_Config = BK_ThisBook.Worksheets("�ݒ�")
'  Set LadexSh_Notice = BK_ThisBook.Worksheets("Notice")
'  Set LadexSh_Style = BK_ThisBook.Worksheets("Style")
'  Set LadexSh_TestData = BK_ThisBook.Worksheets("testData")
'  Set LadexSh_Ribbon = BK_ThisBook.Worksheets("Ribbon")
'  Set LadexSh_Favorite = BK_ThisBook.Worksheets("Favorite")
'  Set LadexSh_Stamp = BK_ThisBook.Worksheets("Stamp")
'  Set LadexSh_HighLight = BK_ThisBook.Worksheets("HighLight")
'  Set LadexSh_Help = BK_ThisBook.Worksheets("Help")
'  Set LadexSh_Function = BK_ThisBook.Worksheets("Function")
'  Set LadexSh_InputData = BK_ThisBook.Worksheets("inputData")
 
 
  '�ݒ�l�ǂݍ���--------------------------------
  Set BK_setVal = Nothing
  Set BK_setVal = CreateObject("Scripting.Dictionary")
  
  endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 0 Then
    endLine = 11
  End If
  
  For line = 3 To endLine
    If LadexSh_Config.Range("A" & line) <> "" Then
      BK_setVal.add LadexSh_Config.Range("A" & line).Text, LadexSh_Config.Range("B" & line).Text
    End If
  Next
    
  '���O�o�͐ݒ�----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  LadexDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex"
  logFile = LadexDir & "\log\ExcelMacro.log"
  Set wsh = Nothing
  
  If Application.UserName = "���� ����" Then
    BK_setVal("debugMode") = "develop"
  End If
  
'  LogLevel = Split(BK_setVal("LogLevel"), ".")(0)
  
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function

'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  Const funcName As String = "init.���O��`"
  
  On Error GoTo catchError

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
  For line = 3 To LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
    If LadexSh_Config.Range("A" & line) <> "" Then
      LadexSh_Config.Range("B" & line).Name = LadexSh_Config.Range("A" & line)
    End If
  Next
  
  'Book�p�̐ݒ�
  LadexSh_Config.Range("D3:D" & LadexSh_Config.Cells(Rows.count, 6).End(xlUp).Row).Name = LadexSh_Config.Range("D2")
  

  Exit Function
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function

