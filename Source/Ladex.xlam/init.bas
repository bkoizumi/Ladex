Attribute VB_Name = "init"
Option Explicit

'���[�N�u�b�N�p�ϐ�------------------------------
Public LadexBook            As Workbook
Public targetBook           As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
Public targetSheet          As Worksheet

'�Z���p�ϐ�--------------------------------------
Public targetRange          As Range


'�O���[�o���ϐ�----------------------------------
Public Const thisAppName    As String = "Ladex"
Public Const thisAppVersion As String = "2.0.0.0"
Public Const RelaxTools     As String = "Relaxtools.xlam"
Public Const thisAppPasswd  As String = "Ladex"


Public funcName             As String
Public runFlg               As Boolean
Public G_LogLevel           As Long

Public arrFavCategory()
Public arrCells()

'�v���O���X�o�[�֘A------------------------------
Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
Public PbarCnt              As Long


'���W�X�g���o�^�p�L�[----------------------------
Public Const RegistryKey    As String = "Ladex"
Public RegistrySubKey       As String


'�ݒ�l�ێ�--------------------------------------
Public dicVal               As Object
Public FrmVal               As Object
Public setIni               As Object
Public sampleDataList       As Object
Public resetVal             As String


'�t�@�C��/�f�B���N�g���֘A-----------------------
Public logFile              As String
Public LadexDir             As String
Public AddInDir             As String


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

'���C�ɓ���֘A----------------------------------
Public Const favoriteDebug  As Boolean = False

'�Z���֘A----------------------------------------
Public Const maxColumnWidth As Long = 60
Public Const maxRowHeight   As Long = 200


'�X�^�C���֘A------------------------------------
Public useStyle()           As Variant
Public useStyleVal          As Object





'**************************************************************************************************
' * �ݒ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "init.unsetting"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  '----------------------------------------------
  If flg = True Then
    Call resetGlobalVal
  End If
  
  Set LadexBook = Nothing
  
  '�ݒ�l�ǂݍ���
  Set dicVal = Nothing
  Set FrmVal = Nothing
  Set useStyleVal = Nothing
  
  Set targetSheet = Nothing
  Set targetRange = Nothing
  
  Erase arrFavCategory
  Erase useStyle
  Erase arrCells
  
  logFile = ""
  LadexDir = ""
  

  
  '�����I��--------------------------------------
  Exit Function
  '----------------------------------------------
  
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function resetGlobalVal()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "init.resetGlobalVal"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call Library.showDebugForm(funcName, , "function")
  '----------------------------------------------

  '�ݒ�l�ǂݍ���
  Set dicVal = Nothing

  
  PrgP_Max = 2
  PrgP_Cnt = 0
  PbarCnt = 1
  runFlg = False
  
  '�����I��--------------------------------------
  Exit Function
  '----------------------------------------------
  
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
  Dim tmpRegList
  
  Const funcName As String = "init.setting"
  
  '�����J�n--------------------------------------
'  On Error GoTo catchError
  '----------------------------------------------

  If LadexDir = "" Or dicVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  '���W�X�g���֘A
  RegistrySubKey = "Main"
  
  '�u�b�N�̐ݒ�
  Set LadexBook = ThisWorkbook
  
  '���O�o�͐ݒ�----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  LadexDir = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex"
  logFile = LadexDir & "\log\ExcelMacro.log"
  AddInDir = wsh.SpecialFolders("AppData") & "\Microsoft\AddIns"

  
  Set wsh = Nothing
  Call Library.showDebugForm(funcName, , "function")
  
  If Library.Book�̏�Ԋm�F = True Then
    '�ݒ�l�ǂݍ���--------------------------------
    Set dicVal = Nothing
    Set dicVal = CreateObject("Scripting.Dictionary")
    
    endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
    If endLine = 0 Then
      endLine = 11
    End If
    
    For line = 3 To endLine
      If LadexSh_Config.Range("A" & line) <> "" Then
        dicVal.add LadexSh_Config.Range("A" & line).Text, LadexSh_Config.Range("B" & line).Text
      End If
    Next
    
    
    '���[�U�[�t�H�[������̎󂯎��p--------------
    Set FrmVal = Nothing
    Set FrmVal = CreateObject("Scripting.Dictionary")
    FrmVal.add "commentVal", ""
    
    '���W�X�g���ݒ荀�ڎ擾------------------------
    tmpRegList = GetAllSettings(thisAppName, "Main")
    For line = 0 To UBound(tmpRegList)
      dicVal.add tmpRegList(line, 0), tmpRegList(line, 1)
    Next
    
    G_LogLevel = Split(dicVal("LogLevel"), ".")(0)
    
  Else
    Set FrmVal = Nothing
    Set FrmVal = CreateObject("Scripting.Dictionary")
    FrmVal.add "LogLevel", "5"
    G_LogLevel = 5
  
  End If
  

  
  
  '�����I��--------------------------------------
  Exit Function
  '----------------------------------------------
  
  
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

'==================================================================================================
Function resetsetVal()
  Dim line As Long, endLine As Long
  Dim tmpRegList
  
  Const funcName As String = "init.resetsetVal"
  
  '�����J�n--------------------------------------
  On Error GoTo catchError
  '----------------------------------------------
  
  '�ݒ�l�ǂݍ���--------------------------------
  Set dicVal = Nothing
  Set dicVal = CreateObject("Scripting.Dictionary")
  
  endLine = LadexSh_Config.Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 0 Then
    endLine = 11
  End If
  
  '���W�X�g���ݒ荀�ڎ擾------------------------
  tmpRegList = GetAllSettings(thisAppName, "Main")
  For line = 0 To UBound(tmpRegList)
    dicVal.add tmpRegList(line, 0), tmpRegList(line, 1)
  Next
    
  Exit Function
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  " & Err.Description
End Function
