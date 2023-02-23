Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook             As Workbook
Public targetBook           As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
Public targetSheet          As Worksheet
Public Sh_PARAM             As Worksheet
Public Sh_WBS               As Worksheet
Public sh_Sumally           As Worksheet
Public sh_Option            As Worksheet

'�O���[�o���ϐ�----------------------------------
Public Const thisAppName    As String = "Work Breakdown Structure for Excel"
Public Const thisAppVersion As String = "1.0.0.0"

Public PrgP_Cnt             As Long
Public PrgP_Max             As Long
Public runFlg               As Boolean
Public reCalflg             As Boolean
Public resetCellFlg         As Boolean

'�ݒ�l�ێ�--------------------------------------
Public setVal               As Object
Public FrmVal               As Object
Public getVal               As Object
Public setAssign            As Object

Public resetVal             As String
Public SlctRange            As Range
Public PBarCnt              As Long

Public Const startLine As Long = 7


'�t�@�C��/�f�B���N�g���֘A-----------------------
Public logFile              As String


'�S���ҏ��--------------------------------------
Public lstAssign()          As String



'***********************************************************************************************************************************************
' * �ݒ�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function unsetting(Optional flg As Boolean = True)
  Const funcName As String = "init.unsetting"
  
  If flg = True Then
    Call Library.showDebugForm("PrgP_Cnt", PrgP_Cnt, "debug")
    Call Library.showDebugForm("PrgP_Max", PrgP_Max, "debug")
  End If
  
  
  Set ThisBook = Nothing
  Set targetBook = Nothing
  
  Set targetSheet = Nothing
  Set Sh_PARAM = Nothing
  Set Sh_WBS = Nothing
  Set sh_Sumally = Nothing
  Set sh_Option = Nothing
  
  Set setVal = Nothing
  Set SlctRange = Nothing
  
  logFile = ""
  reCalflg = False
  PBarCnt = 1
  
  If flg = True Then
    PrgP_Cnt = 1
    PrgP_Max = 0
    
    runFlg = False
  End If
  
End Function
'***********************************************************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function setting(Optional reCheckFlg As Boolean = False)
  Dim line As Long
  Const funcName As String = "init.setting"
  
  On Error GoTo catchError
  
  If logFile = "" Or setVal Is Nothing Or reCheckFlg = True Then
    Call init.unsetting(False)
  Else
    Exit Function
  End If

  Set targetBook = ActiveWorkbook
  
  '���[�N�V�[�g���̐ݒ�
  Set Sh_PARAM = targetBook.Worksheets("PARAM")
  Set Sh_WBS = targetBook.Worksheets("WBS")
  
  '���O�o�͐ݒ�----------------------------------
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  logFile = wsh.SpecialFolders("AppData") & "\Bkoizumi\Ladex\log\WBS_ExcelMacro.log"
  Set wsh = Nothing
  
  
  
  '�ݒ�l�ǂݍ���--------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
'  setVal.Add item:="develop", Key:="debugMode"
  setVal.Add item:="5", Key:="LogLevel"
  
  endLine = Sh_PARAM.Cells(Rows.count, 1).End(xlUp).Row
  On Error Resume Next
  For line = 2 To endLine
    If Sh_PARAM.Range("A" & line) <> "" Then
      setVal.Add Sh_PARAM.Range("A" & line).Text, Sh_PARAM.Range("B" & line).Text
    End If
  Next
'  On Error GoTo catchError
  
  
'  Call WBS_Option.�ݒ�V�[�g�R�s�[("forAddin")
  Set sh_Option = ActiveWorkbook.Worksheets("Option")
  
  
  endLine = sh_Option.Cells(Rows.count, 1).End(xlUp).Row
  For line = 3 To endLine
    If sh_Option.Range("A" & line) <> "" Then
      setVal.Add sh_Option.Range("A" & line).Text, sh_Option.Range("B" & line).Text
    End If
  Next
  

  '�S���ғǂݍ���--------------------------------
  Set setAssign = Nothing
  Set setAssign = CreateObject("Scripting.Dictionary")
  
  endLine = sh_Option.Cells(Rows.count, 11).End(xlUp).Row
  On Error Resume Next
  For line = 4 To endLine
    If sh_Option.Range("K" & line) <> "" Then
      setAssign.Add sh_Option.Range("K" & line).Text, sh_Option.Range("K" & line).Interior.Color
    End If
  Next



  

  
  
  
  Exit Function
  
'�G���[������=====================================================================================
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
  logFile = ""
End Function

'**************************************************************************************************
' * �x���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHollyday(chkDate As Date, HollydayName As String)
  Dim line As Long, endLine As Long
  Dim strFilMessage() As Date
  
  '�x������
  Call GetHollyday(CDate(chkDate), HollydayName)
  
  '�y���𔻒�
  If HollydayName = "" Then
    If Weekday(chkDate) = vbSunday Then
      HollydayName = "Sunday"
    ElseIf Weekday(chkDate) = vbSaturday Then
      HollydayName = "Saturday"
    End If
  End If
  
  
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

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To setSheet.Range("B4")
    If setSheet.Range("A" & line) <> "" Then
      setSheet.Range(setVal("cell_LevelInfo") & line).Name = setSheet.Range("A" & line)
    End If
  Next
  
  '�V���[�g�J�b�g�L�[�̐ݒ�
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).Row
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_ShortcutFuncName") & line) <> "" Then
      setSheet.Range(setVal("cell_ShortcutKey") & line).Name = setSheet.Range(setVal("cell_ShortcutFuncName") & line)
    End If
  Next
  
  
  endLine = setSheet.Cells(Rows.count, 11).End(xlUp).Row
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "�S����"

  endLine = setSheet.Cells(Rows.count, 17).End(xlUp).Row
  setSheet.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "�x�����X�g"

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'***********************************************************************************************************************************************
' * �V�[�g�̕\��/��\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function noDispSheet()

  Call init.setting
  'Worksheets("Help").Visible = xlSheetVeryHidden
  Worksheets("Tmp").Visible = xlSheetVeryHidden
  Worksheets("Notice").Visible = xlSheetVeryHidden
'  Worksheets("�ݒ�").Visible = xlSheetVeryHidden
  Worksheets("�T���v��").Visible = xlSheetVeryHidden
  Worksheets(TeamsPlannerSheetName).Visible = xlSheetVeryHidden
  
  Worksheets(mainSheetName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("�ݒ�").Visible = True
  Worksheets("�T���v��").Visible = True
  
  Worksheets(TeamsPlannerSheetName).Visible = True
  Worksheets(mainSheetName).Visible = True
  
  Worksheets(mainSheetName).Select
  
End Function





































