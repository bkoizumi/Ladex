Attribute VB_Name = "Library"
'**************************************************************************************************
' * �Q�Ɛݒ�A�萔�錾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' ���p����Q�Ɛݒ�܂Ƃ�
' Microsoft Office 14.0 Object Library
' Microsoft DAO 3.6 Objects Library
' Microsoft Scripting Runtime (WSH, FileSystemObject)
' Microsoft ActiveX Data Objects 2.8 Library
' UIAutomationClient

' Windows API�̗��p--------------------------------------------------------------------------------
' �f�B�X�v���C�̉𑜓x�擾�p
' Sleep�֐��̗��p
' �N���b�v�{�[�h�֐��̗��p
#If VBA7 And Win64 Then
  '�f�B�X�v���C�̉𑜓x�擾�p
  Private Declare PtrSafe Function getSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'Sleep�֐��̗��p
  Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As LongPtr)

  '�N���b�v�{�[�h�֘A
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

#Else
  '�f�B�X�v���C�̉𑜓x�擾�p
  Private Declare Function getSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'Sleep�֐��̗��p
  Private Declare Function Sleep Lib "kernel32" (ByVal ms As Long)

  '�N���b�v�{�[�h�֘A
  Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function EmptyClipboard Lib "user32" () As Long


  'Shell�֐��ŋN�������v���O�����̏I����҂�
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Private Const PROCESS_QUERY_INFORMATION = &H400&
  Private Const STILL_ACTIVE = &H103&

#End If



'���[�N�u�b�N�p�ϐ�------------------------------
'���[�N�V�[�g�p�ϐ�------------------------------
'�O���[�o���ϐ�----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

'�A�N�e�B�u�Z���̎擾
Dim SelectionCell As String
Dim SelectionSheet As String

' PC�AOffice���̏��擾�p�A�z�z��
Public MachineInfo As Object

' Selenium�p�ݒ�
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

'Public Const JapaneseCharacters = "�����������������������������������ĂƂȂɂʂ˂̂͂Ђӂւق܂݂ނ߂�������������񂪂����������������������Âłǂ΂тԂׂڂς҂Ղ؂�"
'Public Const JapaneseCharactersCommonUse = "�J�w����щ�⋞���o�m�����X�����������ψ�j�݋��K�n�g�����Ґ̎�󏊒���g�\�������������a�p�ʉ芯�G�����a�ō��Q����������I�T�ŔO�{�@�q��Չ����͋������Ȏ}�ɏq���������Ŕ�񕐉����g���č���@���S�����͓��q���󖇈ˉ���F���������h���������˔t�������ޕ|���b�Ή�����"
'Public Const MachineDependentCharacters = "�@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S�T�U�V�W�X�Y�Z�[�\�]�_�\�]�^�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�z�{"


Public ThisBook As Workbook


'**************************************************************************************************
' * �A�h�I�������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function addinClose()
  Workbooks(ThisWorkbook.Name).Close
End Function


'**************************************************************************************************
' * �G���[���̏���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function errorHandle(FuncName As String, ByRef objErr As Object)

  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  Message = FuncName & vbCrLf & objErr.Description

  '�����F�����b
  Application.Speech.Speak Text:="�G���[���������܂���", SpeakAsync:=True
  Message = Application.WorksheetFunction.VLookup(objErr.Number, noticeCodeSheet.Range("A2:B" & endLine), 2, False)

  Call MsgBox(Message, vbCritical)
  Call endScript
  Call Ctl_ProgressBar.showEnd

  Call outputLog(runTime, objErr.Number & vbTab & objErr.Description)
End Function


'**************************************************************************************************
' * ��ʕ`�ʐ���J�n
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()
  On Error Resume Next
  '�A�N�e�B�u�Z���̎擾
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
    SelectionSheet = ActiveWorkbook.ActiveSheet.Name
  End If

  '�}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = False

  '�}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = False

  '�}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationManual

  '�}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
  'Application.Interactive = False

  '�}�N�����쒆�̓}�E�X�J�[�\�����u�����v�v�ɂ���
  'Application.Cursor = xlWait

  '�m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = False

End Function


'**************************************************************************************************
' * ��ʕ`�ʐ���I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional reCalflg As Boolean = False, Optional flg As Boolean = False)
  On Error Resume Next

  '�����I�ɍČv�Z������
  If reCalflg = True Then
    Application.CalculateFull
  End If

 '�A�N�e�B�u�Z���̑I��
  If SelectionCell <> "" And flg = True Then
    ActiveWorkbook.Worksheets(SelectionSheet).Select
    ActiveWorkbook.Range(SelectionCell).Select
  End If
  Call unsetClipboard

  '�}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = True

  '�}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = True

  '�}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationAutomatic

  '�}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
  'Application.Interactive = True

  '�}�N������I����̓}�E�X�J�[�\�����u�f�t�H���g�v�ɂ��ǂ�
  Application.Cursor = xlDefault

  '�}�N������I����̓X�e�[�^�X�o�[���u�f�t�H���g�v�ɂ��ǂ�
  Application.StatusBar = False

  '�m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = True
End Function


'**************************************************************************************************
' * �V�[�g�̑��݊m�F
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkSheetExists(sheetName) As Boolean

  Dim tempSheet As Object
  Dim result As Boolean

  result = False
  For Each tempSheet In Sheets
    If LCase(sheetName) = LCase(tempSheet.Name) Then
      result = True
      Exit For
    End If
  Next
  chkSheetExists = result
End Function


'**************************************************************************************************
' * ���������܂őҋ@
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShellEnd(ProcessID As Long)
  Dim hProcess As Long
  Dim EndCode As Long
  Dim EndRet   As Long

  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, ProcessID)
  Do
    EndRet = GetExitCodeProcess(hProcess, EndCode)
    DoEvents
  Loop While (EndCode = STILL_ACTIVE)
  EndRet = CloseHandle(hProcess)
End Function


'**************************************************************************************************
' * �I�[�g�V�F�C�v�̑��݊m�F
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShapeName(ShapeName As String) As Boolean

  Dim objShp As Shape
  Dim result As Boolean

  result = False
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name = ShapeName Then
      result = True
      Exit For
    End If
  Next
  chkShapeName = result
End Function


'**************************************************************************************************
' * ���O�V�[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(chkSheetName As String) As Boolean

 Dim result As Boolean
  Dim sheetName As Variant

  For Each sheetName In Range("ExcludeSheet")
    If sheetName = chkSheetName Then
      result = True
      Exit For
    Else
      result = False
    End If
  Next
  chkExcludeSheet = result
End Function


'**************************************************************************************************
' * �z�񂪋󂩂ǂ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
 Function chkArrayEmpty(arrayTmp As Variant) As Boolean

  On Error GoTo catchError

  If UBound(arrayTmp) >= 0 Then
    chkArrayEmpty = False
  Else
    chkArrayEmpty = True
  End If

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  chkArrayEmpty = True

End Function

'**************************************************************************************************
' * �u�b�N���J����Ă��邩�`�F�b�N
' *
' * @Link https://www.moug.net/tech/exvba/0060042.html
'**************************************************************************************************
Function chkBookOpened(chkFile) As Boolean

  Dim myChkBook As Workbook
  On Error Resume Next

  Set myChkBook = Workbooks(chkFile)

  If Err.Number > 0 Then
    chkBookOpened = False
  Else
    chkBookOpened = True
  End If
End Function


'**************************************************************************************************
' * �w�b�_�[�`�F�b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHeader(baseNameArray As Variant, chkNameArray As Variant)
  Dim errMeg As String


On Error GoTo catchError
  errMeg = ""

  If UBound(baseNameArray) <> UBound(chkNameArray) Then
    errMeg = "�����قȂ�܂��B"
    errMeg = errMeg & vbNewLine & UBound(baseNameArray) & "<=>" & UBound(chkNameArray) & vbNewLine
  Else
    For i = LBound(baseNameArray) To UBound(baseNameArray)
      If baseNameArray(i) <> chkNameArray(i) Then
        errMeg = errMeg & vbNewLine & i & ":" & baseNameArray(i) & "<=>" & chkNameArray(i)
      End If
    Next
  End If

  chkHeader = errMeg

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  chkHeader = "�G���[���������܂���"

End Function

'**************************************************************************************************
' * �t�@�C���̑��݊m�F
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkFileExists(targetPath As String)
  Dim FSO As Object

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FileExists(targetPath) Then
    chkFileExists = True
  Else
    chkFileExists = False
  End If
  Set FSO = Nothing

End Function


'**************************************************************************************************
' * �f�B���N�g���̑��݊m�F
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkDirExists(targetPath As String)
  Dim FSO As Object

  Set FSO = CreateObject("Scripting.FileSystemObject")

  If FSO.FolderExists(targetPath) Then
    chkDirExists = True
  Else
    chkDirExists = False
  End If
  Set FSO = Nothing

End Function


'**************************************************************************************************
' * �Œ蒷������ɕϊ�
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFixedLength(strTarget As String, lengs As Long, addString As String) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  Do While LenB(strTarget) <= lengs
    strTarget = strTarget & addString
  Loop
  convFixedLength = strTarget
End Function


'**************************************************************************************************
' * �L�������P�[�X���X�l�[�N�P�[�X�ɕϊ�
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function covCamelToSnake(ByVal val As String, Optional ByVal isUpper As Boolean = False) As String
  Dim ret As String
  Dim i      As Long, Length As Long

  Length = Len(val)

  For i = 1 To Length
    If UCase(Mid(val, i, 1)) = Mid(val, i, 1) Then
      If i = 1 Then
        ret = ret & Mid(val, i, 1)
      ElseIf i > 1 And UCase(Mid(val, i - 1, 1)) = Mid(val, i - 1, 1) Then
        ret = ret & Mid(val, i, 1)
      Else
        ret = ret & "_" & Mid(val, i, 1)
      End If
    Else
      ret = ret & Mid(val, i, 1)
    End If
  Next

  If isUpper Then
    covCamelToSnake = UCase(ret)
  Else
    covCamelToSnake = LCase(ret)
  End If
End Function


'**************************************************************************************************
' * �X�l�[�N�P�[�X���L�������P�[�X�ɕϊ�
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function convSnakeToCamel(ByVal val As String, Optional ByVal isFirstUpper As Boolean = False) As String

  Dim ret As String
  Dim i   As Long
  Dim snakeSplit As Variant

  snakeSplit = Split(val, "_")

  For i = LBound(snakeSplit) To UBound(snakeSplit)
    ret = ret & UCase(Mid(snakeSplit(i), 1, 1)) & Mid(snakeSplit(i), 2, Len(snakeSplit(i)))
  Next

  If isFirstUpper Then
    convSnakeToCamel = ret
  Else
    convSnakeToCamel = LCase(Mid(ret, 1, 1)) & Mid(ret, 2, Len(ret))
  End If
End Function


'**************************************************************************************************
' * ���p�̃J�^�J�i��S�p�̃J�^�J�i�ɕϊ�����(�������p�����͔��p�ɂ���)
' *
' * @link   http://officetanaka.net/excel/function/tips/tips45.htm
'**************************************************************************************************
Function convHan2Zen(Text As String) As String
  Dim i As Long, buf As String

  Dim c As Range
  Dim rData As Variant, ansData As Variant

  For i = 1 To Len(Text)
    DoEvents
    rData = StrConv(Text, vbWide)
    If Mid(rData, i, 1) Like "[�`-��]" Or Mid(rData, i, 1) Like "[�O-�X]" Or Mid(rData, i, 1) Like "[�|�I�i�j�^]" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
      ansData = ansData & Mid(rData, i, 1)
    End If
  Next i
  convHan2Zen = ansData
End Function


'**************************************************************************************************
' * �p�C�v���J���}�ɕϊ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convPipe2Comma(strText As String) As String

  Dim covString As String

  tmp = Split(strText, "|")
  covString = ""
  For i = 0 To UBound(tmp)
    If i = 0 Then
      covString = tmp(i)
    Else
      covString = covString & "," & tmp(i)
    End If
  Next i
  convPipe2Comma = covString

End Function


'**************************************************************************************************
' * Base64�G���R�[�h(�t�@�C��)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" '������
  On Error Resume Next
  Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
  With CreateObject("ADODB.Stream")
    .Type = adTypeBinary
    .Open
    .LoadFromFile filePath
    elm.dataType = "bin.base64"
    elm.nodeTypedValue = .Read(adReadAll)
    ret = elm.Text
    .Close
  End With
  On Error GoTo 0
  convBase64EncodeForFile = ret
End Function


'**************************************************************************************************
' * Base64�G���R�[�h(������)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String

  Dim ret As String
  Dim d() As Byte

  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" '������
  On Error Resume Next
  With CreateObject("ADODB.Stream")
    .Open
    .Type = adTypeText
    .Charset = "UTF-8"
    .WriteText str
    .Position = 0
    .Type = adTypeBinary
    .Position = 3
    d = .Read()
    .Close
  End With
  With CreateObject("MSXML2.DOMDocument").createElement("base64")
    .dataType = "bin.base64"
    .nodeTypedValue = d
    ret = .Text
  End With
  On Error GoTo 0
  convBase64EncodeForString = ret
End Function


'**************************************************************************************************
' * URL-safe Base64�G���R�[�h
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLSafeBase64Encode(ByVal str As String) As String

  str = convBase64EncodeForString(str)
  str = Replace(str, "+", "-")
  str = Replace(str, "/", "_")

  convURLSafeBase64Encode = str
End Function


'**************************************************************************************************
' * URL�G���R�[�h
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLEncode(ByVal str As String) As String

  With CreateObject("ScriptControl")
    .Language = "JScript"
    EncodeURL = .codeobject.encodeURIComponent(str)
  End With
End Function


'**************************************************************************************************
' * �擪�P�����ڂ�啶����
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFirstCharConvert(ByVal strTarget As String) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  strFirst = UCase(Left$(strTarget, 1))
  strExceptFirst = Mid$(strTarget, 2, Len(strTarget))
  convFirstCharConvert = strFirst & strExceptFirst
End Function


'**************************************************************************************************
' * ������̍�������w�蕶�����폜����֐�
' *
' * @Link   https://vbabeginner.net/vba�ŕ�����̉E���⍶������w�蕶�����폜����/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
  Dim iLen    As Long

  '������ł͂Ȃ��ꍇ
  If VarType(s) <> vbString Then
      cutLeft = s & "������ł͂Ȃ�"
      Exit Function
  End If

  iLen = Len(s)

  '�����񒷂��w�蕶�������傫���ꍇ
  If iLen < i Then
      cutLeft = s & "�����񒷂��w�蕶�������傫��"
      Exit Function
  End If

  cutLeft = Right(s, iLen - i)
End Function


'**************************************************************************************************
' * ������̉E������w�蕶�����폜����֐�
' *
' * @Link   https://vbabeginner.net/vba�ŕ�����̉E���⍶������w�蕶�����폜����/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
  Dim iLen    As Long

  If VarType(s) <> vbString Then
    cutRight = s & "������ł͂Ȃ�"
    Exit Function
  End If

  iLen = Len(s)

  '�����񒷂��w�蕶�������傫���ꍇ
  If iLen < i Then
    cutRight = s & "�����񒷂��w�蕶�������傫��"
    Exit Function
  End If

  cutRight = Left(s, iLen - i)
End Function


'**************************************************************************************************
' * �A�����s�̍폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delMultipleLine(targetValue As String)
  With CreateObject("VBScript.RegExp")
    .Global = True
    .Pattern = "(\r\n)+"
    combineMultipleLine = .Replace(targetValue, vbCrLf)
  End With
End Function

'**************************************************************************************************
' * �V�[�g�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delSheetData(Optional line As Long)

  If line <> 0 Then
    Rows(line & ":" & Rows.count).delete Shift:=xlUp
    Rows(line & ":" & Rows.count).Select
    Rows(line & ":" & Rows.count).NumberFormatLocal = "G/�W��"
    Rows(line & ":" & Rows.count).Style = "Normal"
  Else
    Cells.delete Shift:=xlUp
    Cells.NumberFormatLocal = "G/�W��"
    Cells.Style = "Normal"
  End If
  DoEvents

  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function

'**************************************************************************************************
' * �Z�����̉��s�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delCellLinefeed(val As String)
  Dim stringVal As Variant
  Dim retVal As String
  Dim count As Integer

  retVal = ""
  count = 0
  For Each stringVal In Split(val, vbLf)
    If stringVal <> "" And count <= 1 Then
      retVal = retVal & stringVal & vbLf
      count = 0
    Else
      count = count + 1
    End If
  Next
  delCellLinefeed = retVal
End Function

'**************************************************************************************************
' * �I��͈͂̉摜�폜
' *
' * @Link https://www.relief.jp/docs/018407.html
'**************************************************************************************************
Function delImage()
  Dim Rng As Range
  Dim shp As Shape

  If TypeName(Selection) <> "Range" Then
    Exit Function
  End If

  For Each shp In ActiveSheet.Shapes
    Set Rng = Range(shp.TopLeftCell, shp.BottomRightCell)

    If Not (Intersect(Rng, Selection) Is Nothing) Then
      shp.delete
    End If
  Next
End Function


'**************************************************************************************************
' * �Z���̖��̐ݒ�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delVisibleNames()
  Dim Name As Object

  On Error Resume Next

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.delete
    End If
  Next

End Function


'**************************************************************************************************
' * �e�[�u���f�[�^�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delTableData()
  Dim endLine As Long

  On Error Resume Next

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  Rows("3:" & endLine).Select
  Selection.delete Shift:=xlUp

  Rows("2:3").Select
  Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents

  Cells.Select
  Selection.NumberFormatLocal = "G/�W��"

  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * PC�AOffice���̏��擾
' * �A�z�z��𗘗p���Ă���̂ŁAMicrosoft Scripting Runtime���K�{
' * MachineInfo.Item ("Excel") �ŌĂяo��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object

  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  ' OS�̃o�[�W�����擾-----------------------------------------------------------------------------
  Select Case Application.OperatingSystem

    Case "Windows (64-bit) NT 6.01"
        MachineInfo.add "OS", "Windows7-64"

    Case "Windows (32-bit) NT 6.01"
        MachineInfo.add "OS", "Windows7-32"

    Case "Windows (32-bit) NT 5.01"
        MachineInfo.add "OS", "WindowsXP-32"

    Case "Windows (64-bit) NT 5.01"
        MachineInfo.add "OS", "WindowsXP-64"

    Case Else
       MachineInfo.add "OS", Application.OperatingSystem
  End Select

  ' Excel�̃o�[�W�����擾--------------------------------------------------------------------------
  Select Case Application.Version
    Case "16.0"
        MachineInfo.add "Excel", "2016"
    Case "14.0"
        MachineInfo.add "Excel", "2010"
    Case "12.0"
        MachineInfo.add "Excel", "2007"
    Case "11.0"
        MachineInfo.add "Excel", "2003"
    Case "10.0"
        MachineInfo.add "Excel", "2002"
    Case "9.0"
        MachineInfo.add "Excel", "2000"
    Case Else
       MachineInfo.add "Excel", Application.Version
  End Select

  'PC�̏��----------------------------------------------------------------------------------------
  MachineInfo.add "UserName", WshNetworkObject.UserName
  MachineInfo.add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.add "UserDomain", WshNetworkObject.UserDomain

  '��ʂ̉𑜓x���擾------------------------------------------------------------------------------
'  MachineInfo.Add "monitors", getSystemMetrics(80)
'  MachineInfo.Add "displayX", getSystemMetrics(0)
'  MachineInfo.Add "displayY", getSystemMetrics(1)
'  MachineInfo.Add "displayVirtualX", getSystemMetrics(78)
'  MachineInfo.Add "displayVirtualY", getSystemMetrics(79)
'  MachineInfo.Add "appTop", ActiveWindow.Top
'  MachineInfo.Add "appLeft", ActiveWindow.Left
'  MachineInfo.Add "appWidth", ActiveWindow.Width
'  MachineInfo.Add "appHeight", ActiveWindow.Height
End Function


'**************************************************************************************************
' * �������J�E���g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getByteString(arryColumn As String, Optional line As Long) As Long
  Dim colLineName As Variant
  Dim count As Long

  count = 0
  For Each colLineName In Split(arryColumn, ",")
    If line > 0 Then
      count = count + LenB(Range(colLineName & line).Value)
    Else
      count = count + LenB(Range(colLineName).Value)
    End If
  Next colLineName

  getByteString = count
End Function


'**************************************************************************************************
' * �Z���̍��W�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getCellPosition(Rng As Range, ActvCellTop As Long, ActvCellLeft As Long)

  Dim R1C1Top As Long, R1C1Left As Long




  R1C1Top = ActiveWindow.PointsToScreenPixelsY(0)
  R1C1Left = ActiveWindow.PointsToScreenPixelsX(0)

'  ActvCellTop = ((R1C1Top * DPI / PPI) * (ActiveWindow.Zoom / 100)) + Rng.Top
'  ActvCellLeft = ((R1C1Left * DPI / PPI) * (ActiveWindow.Zoom / 100)) + Rng.Left

  ActvCellTop = (((Rng.Top * (DPI / PPI)) * (ActiveWindow.Zoom / 100)) + R1C1Top) * (PPI / DPI)
  ActvCellLeft = (((Rng.Left * (DPI / PPI)) * (ActiveWindow.Zoom / 100)) + R1C1Left) * (PPI / DPI)





'  If ActvCellLeft <= 0 Then
'    ActvCellLeft = 20
'  End If

  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("R1C1Top �F" & R1C1Top)
  Call Library.showDebugForm("R1C1Left �F" & R1C1Left)
  Call Library.showDebugForm("-------------------------")
  Call Library.showDebugForm("Rng.Address �F" & Rng.Address)
  Call Library.showDebugForm("ActvCellTop �F" & ActvCellTop)
  Call Library.showDebugForm("ActvCellLeft�F" & ActvCellLeft)
End Function


'**************************************************************************************************
' * �񖼂����ԍ������߂�
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long

  getColumnNo = Range(targetCell & ":" & targetCell).Column
End Function


'**************************************************************************************************
' * ��ԍ�����񖼂����߂�
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String

  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function

'**************************************************************************************************
' * �J���[�p���b�g��\�����A�F�R�[�h���擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getColor(colorValue As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Call getRGB(colorValue, Red, Green, Blue)
  Application.Dialogs(xlDialogEditColor).Show 10, Red, Green, Blue

  setColorValue = ActiveWorkbook.Colors(10)
  If setColorValue = False Then
    setColorValue = colorValue
  End If

  getColor = setColorValue

End Function

'**************************************************************************************************
' * �t�H���g�_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFont(FontName As String, fontSize As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long

  Application.Dialogs(xlDialogActiveCellFont).Show FontName, "���M�����[", fontSize


End Function


'**************************************************************************************************
' * IndentLevel�l�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getIndentLevel(targetRange As Range)
  Dim thisTargetSheet As Worksheet

  Application.Volatile

  If targetRange = "" Then
    getIndentLevel = ""
  Else
    getIndentLevel = targetRange.IndentLevel + 1
  End If
End Function


'**************************************************************************************************
' * RGB�l�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function





'**************************************************************************************************
' * �f�B���N�g���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getDirPath(CurrentDirectory As String, Optional title As String)

  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = CurrentDirectory & "\"
    .AllowMultiSelect = False

    If title <> "" Then
      .title = title & "�̏ꏊ��I�����Ă�������"
    Else
      .title = "�t�H���_�[��I�����Ă�������"
    End If

    If .Show = True Then
      getDirPath = .SelectedItems(1)
    Else
      getDirPath = ""
    End If
  End With
End Function


'**************************************************************************************************
' * �t�@�C���ۑ��_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)

  Dim filePath As String
  Dim result As Long

  Dim fileName As Variant

  fileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excel�t�@�C��,*.xlsx,Excel2003�ȑO,*.xls,Excel�}�N���u�b�N,*.xlsm,���ׂẴt�@�C��, *.*", _
      FilterIndex:=FileTypeNo)

  If fileName <> "False" Then
    getSaveFilePath = filePath
  Else
    getSaveFilePath = ""
  End If
End Function

'**************************************************************************************************
' * �t�@�C���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath As String
  Dim result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' �t�@�C���̎�ނ�ݒ�
    .Filters.Clear
    .Filters.add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm"
    .Filters.add "CSV�t�@�C��", "*.csv"
    .Filters.add "SQL�t�@�C��", "*.sql"
    .Filters.add "�e�L�X�g�t�@�C��", "*.txt"
    .Filters.add "JSON�t�@�C��", "*.json"
    .Filters.add "Accesss�f�[�^�x�[�X", "*.mdb"
    .Filters.add "���ׂẴt�@�C��", "*.*"

    .FilterIndex = FileTypeNo

    '�\������t�H���_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '�\���`���̐ݒ�
    .InitialView = msoFileDialogViewWebView

    '�_�C�A���O �{�b�N�X�̃^�C�g���ݒ�
    .title = title


    If .Show = -1 Then
      filePath = .SelectedItems(1)
    Else
      filePath = ""
    End If
  End With

  getFilePath = filePath

End Function


'**************************************************************************************************
' * �����t�@�C���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath() As Variant
  Dim result As Long

  With Application.FileDialog(msoFileDialogFilePicker)
    '�����I��������
    .AllowMultiSelect = True

    ' �t�@�C���̎�ނ�ݒ�
    .Filters.Clear
    .Filters.add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm"
    .Filters.add "CSV�t�@�C��", "*.csv"
    .Filters.add "SQL�t�@�C��", "*.sql"
    .Filters.add "�e�L�X�g�t�@�C��", "*.txt"
    .Filters.add "JSON�t�@�C��", "*.json"
    .Filters.add "Accesss�f�[�^�x�[�X", "*.mdb"
    .Filters.add "���ׂẴt�@�C��", "*.*"

    .FilterIndex = FileTypeNo

    '�\������t�H���_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '�\���`���̐ݒ�
    .InitialView = msoFileDialogViewWebView

    '�_�C�A���O �{�b�N�X�̃^�C�g���ݒ�
    .title = title


    If .Show = -1 Then
      ReDim Preserve filePath(.SelectedItems.count - 1)
      For i = 1 To .SelectedItems.count
        filePath(i - 1) = .SelectedItems(i)
      Next i
    Else
      ReDim Preserve filePath(0)
      filePath(0) = ""
    End If
  End With

  getFilesPath = filePath

End Function

'**************************************************************************************************
' * �f�B���N�g�����̃t�@�C���ꗗ�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileList(Path As String, fileName As String)
  Dim f As Object, cnt As Long
  Dim list() As String

  cnt = 0
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(Path).Files
      If f.Name Like fileName Then
        ReDim Preserve list(cnt)
        list(cnt) = f.Name
        cnt = cnt + 1
      End If
    Next f
  End With

  getFileList = list
End Function


'**************************************************************************************************
' * �t�@�C�����擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileInfo(targetFilePath As String, Optional fileInfo As Object, Optional getType As String)
  Dim FSO As Object
  Dim fileObject As Object
  Dim sp As Shape

  Set FSO = CreateObject("Scripting.FileSystemObject")

  Set fileInfo = Nothing
  Set fileInfo = CreateObject("Scripting.Dictionary")



  '�쐬����
  fileInfo.add "create_at", FSO.GetFile(targetFilePath).DateCreated

  '�X�V����
  fileInfo.add "update_at", FSO.GetFile(targetFilePath).DateLastModified

  '�t�@�C���T�C�Y
  fileInfo.add "size", FSO.GetFile(targetFilePath).Size

  '�t�@�C���̎��
  fileInfo.add "type", FSO.GetFile(targetFilePath).Type

  '�g���q
  fileInfo.add "extension", FSO.GetExtensionName(targetFilePath)

  '�t�@�C�������݂���t�H���_
  fileInfo.add "CurrentDir", FSO.GetFile(targetFilePath).ParentFolder

  Select Case FSO.GetExtensionName(targetFilePath)
    Case "mp4"

    Case "png"
    Set sp = ActiveSheet.Shapes.AddPicture( _
              fileName:=targetFilePath, _
              LinkToFile:=False, _
              SaveWithDocument:=True, _
              Left:=0, _
              Top:=0, _
              Width:=0, _
              Height:=0 _
              )
    With sp
      .LockAspectRatio = msoTrue
      .ScaleHeight 1, msoTrue
      .ScaleWidth 1, msoTrue

      fileInfo.add "width", CLng(.Width * 4 / 3)
      fileInfo.add "height", CLng(.Height * 4 / 3)
      .delete
    End With
    
    Case "bmp", "jpg", "gif", "emf", "ico", "rle", "wmf"
      Set fileObject = LoadPicture(targetFilePath)
      fileInfo.add "width", fileObject.Width
      fileInfo.add "height", fileObject.Height

      Set fileObject = Nothing

    Case Else
  End Select
  
  Set FSO = Nothing
  
  If getType <> "" Then
    getFileInfo = fileInfo(getType)
    Set fileInfo = Nothing
  End If
  
End Function


'**************************************************************************************************
' * �w��o�C�g���̌Œ蒷�f�[�^�쐬(�����񏈗�)
' *
' * @Link http://www.asahi-net.or.jp/~ef2o-inue/vba_o/function05_110_055.html
'**************************************************************************************************
Function getFixlng(strInText As String, lngFixBytes As Long) As String
    Dim lngKeta As Long
    Dim lngByte As Long, lngByte2 As Long, lngByte3 As Long
    Dim ix As Long
    Dim intCHAR As Long
    Dim strOutText As String

    lngKeta = Len(strInText)
    strOutText = strInText
    ' �o�C�g������
    For ix = 1 To lngKeta
        ' 1���������p/�S�p�𔻒f
        intCHAR = Asc(Mid(strInText, ix, 1))
        ' �S�p�Ɣ��f�����ꍇ�̓o�C�g����1��������
        If ((intCHAR < 0) Or (intCHAR > 255)) Then
            lngByte2 = 2        ' �S�p
        Else
            lngByte2 = 1        ' ���p
        End If
        ' �����ӂꔻ��(�E�؂�̂�)
        lngByte3 = lngByte + lngByte2
        If lngByte3 >= lngFixBytes Then
            If lngByte3 > lngFixBytes Then
                strOutText = Left(strInText, ix - 1)
            Else
                strOutText = Left(strInText, ix)
                lngByte = lngByte3
            End If
            Exit For
        End If
        lngByte = lngByte3
    Next ix
    ' ���s������(�󔒕����ǉ�)
    If lngByte < lngFixBytes Then
        strOutText = strOutText & Space(lngFixBytes - lngByte)
    End If
    getFixlng = strOutText
End Function


'**************************************************************************************************
' * �V�[�g���X�g�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSheetList(columnName As String)

  Dim i As Long
  Dim sheetName As Object

  i = 3
  If columnName = "" Then
    columnName = "E"
  End If

  On Error GoTo GetSheetListError:
  Call startScript

  '���ݒ�l�̃N���A
  Worksheets("�ݒ�").Range(columnName & "3:" & columnName & "100").Select
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle = xlNone
  Selection.Borders(xlEdgeTop).LineStyle = xlNone
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  Selection.Borders(xlInsideVertical).LineStyle = xlNone
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With

  For Each sheetName In ActiveWorkbook.Sheets

    '�V�[�g���̐ݒ�
    Worksheets("�ݒ�").Range(columnName & i).Select
    Worksheets("�ݒ�").Range(columnName & i) = sheetName.Name

    ' �Z���̔w�i�F����
    With Worksheets("�ݒ�").Range(columnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' �V�[�g�F�Ɠ����F���Z���ɐݒ�
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("�ݒ�").Range(columnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    '�r���̐ݒ�
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    i = i + 1
  Next

  Worksheets("�ݒ�").Range(columnName & "3").Select
  Call endScript
  Exit Function
'==================================================================================================
'�G���[�������̏���
'==================================================================================================
GetSheetListError:

  ' ��ʕ`�ʐ���I��
  Call endScript
  Call errorHandle("�V�[�g���X�g�擾", Err)

End Function


'**************************************************************************************************
' * �I���Z���̊g��\���ďo
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionForm(Text As String, SetSelectTargetRows As String)
  With ExpansionForm
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 10)
    .Left = Application.Left + (ActiveWindow.Height / 5)
    .TextBox = Text
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    .Caption = SetSelectTargetRows
  End With
  ExpansionForm.Show vbModeless
End Function


'**************************************************************************************************
' * �f�o�b�O�p��ʕ\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(meg1 As String, Optional meg2 As String)
  Dim runTime As Date
  Dim StartUpPosition As Long

'  On Error GoTo catchError

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  If BK_setVal("debugMode") = "none" Then
    Exit Function
  End If

'  If StopTime <> 0 Then
'    meg1 = meg1 & vbNewLine & "<�������ԁF" & StopTime & ">"
'  End If
  meg1 = Replace(meg1, vbNewLine, " ")


  Select Case BK_setVal("debugMode")
    Case "file"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If
      GoTo label_end

    Case "form"
      GoTo label_showForm

    Case "all"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If
      GoTo label_showForm

    Case "develop"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
        Debug.Print runTime & vbTab & meg1
      End If
      'GoTo label_showForm
      GoTo label_end

    Case Else
      Exit Function
  End Select

label_showForm:
  If meg1 Like "�����J�n�F*" Then

    With Frm_debug
      .Caption = "�������"
      .ListBox1.Clear
      .ListBox1.AddItem runTime & vbTab & meg1
    End With
  Else
    With Frm_debug
      .Caption = "�������"
      .ListBox1.AddItem runTime & vbTab & meg1
      .ListBox1.ListIndex = .ListBox1.ListCount - 1
    End With
  End If

  If (Frm_debug.Visible = True) Then
    Frm_debug.StartUpPosition = 0
  Else
    Frm_debug.StartUpPosition = 1
  End If
  Frm_debug.Show vbModeless


label_end:

  DoEvents
  Exit Function

'�G���[������=====================================================================================
catchError:
  Exit Function
End Function

'**************************************************************************************************
' * �������ʒm
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional process As String, Optional runEndflg As Boolean)
  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  On Error GoTo catchError


  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  endLine = BK_sheetNotice.Cells(Rows.count, 1).End(xlUp).Row
  Message = Application.WorksheetFunction.VLookup(Code, BK_sheetNotice.Range("A2:B" & endLine), 2, False)
  Message = Replace(Message, "%%", process)
  If process = "" Then
    Message = Replace(Message, "<>", process)
  End If
  If runEndflg = True Then
    Message = Message & vbNewLine & "�����𒆎~���܂�"
  End If

  If StopTime <> 0 Then
    Message = Message & vbNewLine & "<�������ԁF" & StopTime & ">"
  End If

  If Message <> "" Then
    Message = Replace(Message, "<BR>", vbNewLine)
  End If

  If BK_setVal("debugMode") = "speak" Or BK_setVal("debugMode") = "develop" Or BK_setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=Message, SpeakAsync:=True, SpeakXML:=True
  End If

  Select Case Code
    Case 0 To 399
      Call MsgBox(Message, vbInformation, thisAppName)

    Case 400 To 499
      Call MsgBox(Message, vbCritical, thisAppName)

    Case 500 To 599
      Call MsgBox(Message, vbExclamation, thisAppName)

    Case 999

    Case Else
      Call MsgBox(Message, vbCritical, thisAppName)
  End Select
'  Stop

  '��ʕ`�ʐ���I������
  If runEndflg = True Then
    Call endScript
    Call Ctl_ProgressBar.showEnd
    End
  Else
    Call Library.showDebugForm(Message)
  End If

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call MsgBox(Message, vbCritical, thisAppName)

End Function


'**************************************************************************************************
' * �����_��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setString As String, ByVal setStringCnt As Integer) As String

 For i = 1 To setStringCnt
    '�����W�F�l���[�^��������
    Randomize
    n = Int((Len(setString) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(setString, n, 1)
  Next i

  makeRandomString = str1

End Function

'==================================================================================================
Function makeRandomNo(minNo As Long, maxNo As Long) As String

  Randomize
  makeRandomNo = Int((maxNo - minNo + 1) * Rnd + minNo)

End Function


'==================================================================================================
Function makeRandomDigits(maxCount As Long) As String
  Dim makeVal As String
  Dim tmpVal As String

  For count = 1 To maxCount
    Randomize
    tmpVal = CStr(Int(10 * Rnd))

    If count = 1 And tmpVal = 0 Then
      tmpVal = 1
    End If
    makeVal = makeVal & tmpVal
  Next

  makeRandomDigits = makeVal

End Function
'**************************************************************************************************
' * ���O�o��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function outputLog(runTime As Date, Message As String)
  Dim fileTimestamp As Date

  If chkFileExists(logFile) Then
    fileTimestamp = FileDateTime(logFile)
  Else
      fileTimestamp = DateAdd("d", -1, Date)
  End If

  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    If Format(Date, "yyyymmdd") = Format(fileTimestamp, "yyyymmdd") Then
      .LoadFromFile logFile
      .Position = .Size
    End If
    .WriteText runTime & Message, 1
    .SaveToFile logFile, 2
    .Close
  End With

End Function

'==================================================================================================
Function outputText(Message As String, outputFilePath)

  Open outputFilePath For Output As #1
  Print #1, Message
  Close #1

End Function

'**************************************************************************************************
' * CSV�C���|�[�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @link   https://www.tipsfound.com/vba/18014
'**************************************************************************************************
Function importCsv(filePath As String, Optional readLine As Long, Optional TextFormat As Variant)

  Dim ws As Worksheet
  Dim qt As QueryTable
  Dim count As Long, line As Long, endLine As Long

  endLine = Cells(Rows.count, 1).End(xlUp).Row
  If endLine = 1 Then
    endLine = 1
  Else
    endLine = endLine + 1
  End If

  If readLine < 1 Then
    readLine = 1
  End If

  Set ws = ActiveSheet
  Set qt = ws.QueryTables.add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A" & endLine))
  With qt
    .TextFilePlatform = 932          ' Shift-JIS ���J��
    .TextFileParseType = xlDelimited ' �����ŋ�؂����`��
    .TextFileCommaDelimiter = True   ' ��؂蕶���̓J���}
    .TextFileStartRow = readLine     ' 1 �s�ڂ���ǂݍ���
    .AdjustColumnWidth = False       ' �񕝂������������Ȃ�
    .RefreshStyle = xlOverwriteCells '�㏑�����w��
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' ���p���̎w��

    If IsArray(TextFormat) Then
      .TextFileColumnDataTypes = TextFormat
    End If

    .Refresh
    DoEvents
    .delete
  End With
  Set qt = Nothing
  Set ws = Nothing

  Call Library.startScript
End Function


'**************************************************************************************************
' * Excel�t�@�C���̃C���|�[�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function importXlsx(filePath As String, targeSheet As String, targeArea As String, dictSheet As Worksheet, Optional passWord As String)

  On Error GoTo catchError
  If passWord <> "" Then
    Workbooks.Open fileName:=filePath, ReadOnly:=True, passWord:=passWord
  Else
    Workbooks.Open fileName:=filePath, ReadOnly:=True
  End If

  If Worksheets(targeSheet).Visible = False Then
    Worksheets(targeSheet).Visible = True
  End If
  Sheets(targeSheet).Select

  ActiveWorkbook.Sheets(targeSheet).Rows.Hidden = False
  ActiveWorkbook.Sheets(targeSheet).Columns.Hidden = False

  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

  ActiveWorkbook.Sheets(targeSheet).Range(targeArea).Copy
  dictSheet.Range("A1").PasteSpecial xlPasteValues

  Application.CutCopyMode = False
  ActiveWorkbook.Close SaveChanges:=False
  dictSheet.Range("A1").Select

  DoEvents
  Call Library.startScript

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName, True)
End Function


'**************************************************************************************************
' * MkDir�ŊK�w�̐[���t�H���_�[�����
' *
' * @link http://officetanaka.net/excel/vba/filesystemobject/sample10.htm
'**************************************************************************************************
Function makeDir(fullPath As String)

  If chkDirExists(fullPath) Then
    Exit Function
  End If
  Call chkParentDir(fullPath)
End Function

'==================================================================================================
Private Function chkParentDir(TargetFolder)
  Dim ParentFolder As String, FSO As Object

  On Error GoTo catchError
  Set FSO = CreateObject("Scripting.FileSystemObject")

  ParentFolder = FSO.GetParentFolderName(TargetFolder)
  If Not FSO.FolderExists(ParentFolder) Then
    Call chkParentDir(ParentFolder)
  End If

  FSO.CreateFolder TargetFolder
  Set FSO = Nothing

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "�f�B���N�g���̍쐬�Ɏ��s���܂���" & vbNewLine & Err.Description, True)
End Function


'**************************************************************************************************
' * �p�X���[�h����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String

  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
    '�����W�F�l���[�^��������
    Randomize
    n = Int((Len(halfChar) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(halfChar, n, 1)
  Next i
  makePasswd = str1
End Function


'**************************************************************************************************
' * �n�C���C�g��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setHighLight(ByVal Target As Range)



End Function

'==================================================================================================
Function unsetHighLight()
  Static xRow
  Static xColumn

  pRow = Selection.Row
  pColumn = Selection.Column
  xRow = pRow
  xColumn = pColumn
  If xColumn <> "" Then
    With Columns(xColumn).Interior
      .ColorIndex = xlNone
    End With
    With Rows(xRow).Interior
      .ColorIndex = xlNone
    End With
  End If

End Function



'**************************************************************************************************
' * �����񕪊�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function splitString(targetString As String, separator As String, count As Integer) As String
  Dim tmp As Variant

  If targetString <> "" Then
    tmp = Split(targetString, separator)
    splitString = tmp(count)
  Else
    splitString = ""
  End If
End Function


'**************************************************************************************************
' * �z��̍Ō�ɒǉ�����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setArrayPush(arrName As Variant, str As Variant)
  Dim i As Long

  i = UBound(arrName)
  If i = 0 Then

  Else
    i = i + 1
    ReDim Preserve arrName(i)
  End If
  arrName(i) = str

End Function


'**************************************************************************************************
' * �t�H���g�J���[�ݒ�
' *
' * @Link https://vbabeginner.net/vba�ŃZ���̎w�蕶����̐F�⑾����ύX����/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font     'Font�I�u�W�F�N�g
  Dim i               '����������̃Z���̈ʒu
  Dim iLen            '����������̕�����
  Dim R   As Range    '�Z���͈͂̂P�Z��

  iLen = Len(a_sSearch)
  i = 1

  For Each R In Selection
    Do
      i = InStr(i, R.Value, a_sSearch)
      If (i = 0) Then
        i = 1
        Exit Do
      End If
      Set f = R.Characters(i, iLen).Font
      f.Color = a_lColor
      f.Bold = a_bBold
      i = i + 1
    Loop
  Next
End Function


'**************************************************************************************************
' * ���W�X�g���֘A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function setRegistry(RegistrySubKey As String, RegistryKey As String, setVal As Variant)

  If getRegistry(RegistrySubKey, RegistryKey) <> setVal And RegistryKey <> "" Then
    Call SaveSetting(thisAppName, RegistrySubKey, RegistryKey, setVal)
  End If
End Function

'==================================================================================================
Function getRegistry(RegistrySubKey As String, RegistryKey As String)
  Dim regVal As String

  On Error GoTo catchError

  If RegistryKey <> "" Then
    regVal = GetSetting(thisAppName, RegistrySubKey, RegistryKey)
  End If
  If regVal = "" Then
    getRegistry = 0
  Else
    getRegistry = regVal
  End If

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description)
End Function


'==================================================================================================
Function delRegistry(RegistrySubKey As String, Optional RegistryKey As String)

  Dim regVal As String

  On Error GoTo catchError
  If RegistryKey = "" Then
    Call DeleteSetting(thisAppName, RegistrySubKey)
  Else
    Call DeleteSetting(thisAppName, RegistrySubKey, RegistryKey)
  End If

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
'  Call Library.showNotice(400, Err.Description, True)
End Function


'**************************************************************************************************
' * �Q�Ɛݒ�������ōs��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setReferences(BookType As String)

  On Error GoTo Err_SetReferences:

  'Microsoft Scripting Runtime (Windows Script Host / FileSystemObject)----------------------------
    LibScript = "C:\Windows\System32\scrrun.dll"
    If Dir(LibScript) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibScript)
    Else
      MsgBox ("Microsoft Scripting Runtime�𗘗p�ł��܂���B" & vbLf & "���p�ł��Ȃ��@�\������܂�")
    End If

  'Microsoft ActiveX Data Objects Library 6.1 (ADO)------------------------------------------------
  If BookType = "DataBase" Then
    LibADO = "C:\Program Files\Common Files\System\Ado\msado15.dll"
    If Dir(LibADO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibADO)
    Else
      MsgBox ("Microsoft ActiveX Data Objects�𗘗p�ł��܂���" & vbLf & "���p�ł��Ȃ��@�\������܂�")
    End If

  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
  LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If Dir(LibDAO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
    Else
      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
      If Dir(LibDAO) <> "" Then
        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
      Else
        MsgBox ("Microsoft DAO 3.6 Objects Library�𗘗p�ł��܂���" & vbLf & "DB�ւ̐ڑ��@�\�����p�ł��܂���")
      End If
    End If
  End If

  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
  If BookType = "" Then
    LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If Dir(LibDAO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
    Else
      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
      If Dir(LibDAO) <> "" Then
        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
      Else
        MsgBox ("Microsoft DAO 3.6 Objects Library�𗘗p�ł��܂���" & vbLf & "DB�ւ̐ڑ��@�\�����p�ł��܂���")
      End If
    End If
  End If


Func_Exit:
  Set Ref = Nothing
  Exit Function

Err_SetReferences:
  If Err.Number = 32813 Then
    Resume Next
  ElseIf Err.Number = 1004 Then
    MsgBox ("�uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v�ɕύX���I")
  Else
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
  End If
End Function


'**************************************************************************************************
' * �I���Z���̍s�w�i�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor(setArea As String, DisType As Boolean, SetColor As String)

  Range(setArea).Select

  '�����t���������N���A
  Selection.FormatConditions.delete

  If DisType = False Then
    '�s�����ݒ�
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    '�s�Ɨ�ɐݒ�
    Selection.FormatConditions.add Type:=xlExpression, Formula1:="=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
  End If

  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = SetColor
'    .Interior.TintAndShade = 0
'    .Font.ColorIndex = 1
  End With
  Selection.FormatConditions(1).StopIfTrue = False
End Function


'**************************************************************************************************
' * �V�[�g�̕ی�/�ی����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setProtectSheet()

  ActiveSheet.Protect passWord:=thisAppPasswd, DrawingObjects:=True, Contents:=True, Scenarios:=True
  ActiveSheet.EnableSelection = xlUnlockedCells

End Function

'==================================================================================================
Function unsetProtectSheet()

  ActiveSheet.Unprotect passWord:=thisAppPasswd
End Function


'**************************************************************************************************
' * �ŏ��̃V�[�g��I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setFirstsheet()
  Dim i As Long

  For i = 1 To Sheets.count
    If Sheets(i).Visible = xlSheetVisible Then
      Sheets(i).Select
      Exit Function
    End If
  Next i
End Function


'**************************************************************************************************
' * �t�@�C���S�̂̕�����u��
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(fileName As String, TargetText As String, Optional NewText As String = "")

 Dim FSO         As FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
 Dim Txt         As TextStream       '�e�L�X�g�X�g���[���I�u�W�F�N�g
 Dim buf_strTxt  As String           '�ǂݍ��݃o�b�t�@

 On Error GoTo Func_Err:

 '�I�u�W�F�N�g�쐬
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(fileName, ForReading)

 '�S���ǂݍ���
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '���t�@�C�������l�[�����āA�e���|�����t�@�C���쐬
  Name fileName As fileName & "_"

  '�u������
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  '�����ݗp�e�L�X�g�t�@�C���쐬
   Set Txt = FSO.CreateTextFile(fileName, True)
  '������
  Txt.Write buf_strTxt
  Txt.Close

  '�e���|�����t�@�C�����폜
  FSO.DeleteFile fileName & "_"

'�I������
Func_Exit:
    Set Txt = Nothing
    Set FSO = Nothing
    Exit Function

Func_Err:
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
End Function


'**************************************************************************************************
' * VBA��Excel�̃R�����g���ꊇ�Ŏ����T�C�Y�ɂ��ăJ�b�R�悭����
' *
' * @Link   http://techoh.net/customize-excel-comment-by-vba/
'**************************************************************************************************
Function setComment(Optional BgColorVal, Optional FontVal)
    Dim cl As Range
    Dim count As Long

    count = 0
    For Each cl In Selection
      count = count + 1
      DoEvents
      If Not cl.Comment Is Nothing Then
        With cl.Comment.Shape
          '�T�C�Y�ݒ�
          .TextFrame.AutoSize = True
          .TextFrame.Characters.Font.Size = 9

          '�`����p�ێl�p�`�ɕύX
          .AutoShapeType = msoShapeRectangle

          '�F
          .line.ForeColor.RGB = RGB(128, 128, 128)
          '.Fill.ForeColor.RGB = RGB(240, 240, 240)
          .Fill.ForeColor.RGB = BgColorVal

          '�e ���ߗ� 30%�A�I�t�Z�b�g�� x:1px,y:1px
          .Shadow.Transparency = 0.3
          .Shadow.OffsetX = 1
          .Shadow.OffsetY = 1

          ' ���������A��������
          .TextFrame.Characters.Font.Bold = False
          .TextFrame.HorizontalAlignment = xlLeft

'          .TextFrame.Characters.Font = FontVal

          ' �Z���ɍ��킹�Ĉړ�����
          .Placement = xlMove
        End With
      End If
    Next cl

End Function



'**************************************************************************************************
' * �N���b�v�{�[�h�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetClipboard()
'  OpenClipboard 0
'  EmptyClipboard
'  CloseClipboard
End Function


'**************************************************************************************************
' * �I���Z���̍s�w�i����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLineColor(setArea As String)
  ActiveWorkbook.ActiveSheet.Range(setArea).Select

  '�����t���������N���A
  Selection.FormatConditions.delete
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * �����N����
' *
' * @Link   https://excel-excellent-technics.com/excel-vba-breaklinks-1019
'**************************************************************************************************
Function unsetLink()
  Dim wb          As Workbook
  Dim vntLink     As Variant
  Dim i           As Integer

  Set wb = ActiveWorkbook
  vntLink = wb.LinkSources(xlLinkTypeExcelLinks) '�u�b�N�̒��ɂ��郊���N

  If IsArray(vntLink) Then
    For i = 1 To UBound(vntLink)
      wb.BreakLink vntLink(i), xlLinkTypeExcelLinks '�����N����
    Next i
  End If
End Function


'**************************************************************************************************
' * �X���[�v����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  DoEvents
  'Sleep timeVal

  Application.Wait [Now()] + timeVal / 86400000
  DoEvents
End Function





'**************************************************************************************************
' * �r��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �r��_�N���A(Optional setArea As Range)
  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  End If
End Function

'==================================================================================================
Function �r��_�\(Optional setArea As Range, Optional LineColor As Long)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = xlThin
      .Borders(xlEdgeRight).Weight = xlThin
      .Borders(xlEdgeTop).Weight = xlThin
      .Borders(xlEdgeBottom).Weight = xlThin

      .Borders(xlInsideVertical).Weight = xlThin
      .Borders(xlInsideHorizontal).Weight = xlHairline

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)

        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = xlThin
      .Borders(xlEdgeRight).Weight = xlThin
      .Borders(xlEdgeTop).Weight = xlThin
      .Borders(xlEdgeBottom).Weight = xlThin

      .Borders(xlInsideVertical).Weight = xlThin
      .Borders(xlInsideHorizontal).Weight = xlHairline

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)

        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function



'==================================================================================================
Function �r��_�j��_�͂�(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_�j��_�i�q(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_�j��_���E(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'==================================================================================================
Function �r��_�j��_�㉺(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_�j��_����(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_�j��_����(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With

  End If
End Function





'==================================================================================================
Function �r��_����_�͂�(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_����_�i�q(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_����_���E(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous

'      .Borders(xlEdgeLeft).Weight = WeightVal
'      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous

'      .Borders(xlEdgeLeft).Weight = WeightVal
'      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With

  End If
End Function

'==================================================================================================
Function �r��_����_�㉺(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_����_����(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlContinuous
'      .Borders(xlInsideVertical).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlContinuous
'      .Borders(xlInsideVertical).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_����_����(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
'      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
'      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With

  End If
End Function


'==================================================================================================
Function �r��_��d��_�͂�(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble

'      .Borders(xlEdgeLeft).Weight = WeightVal
'      .Borders(xlEdgeRight).Weight = WeightVal
'      .Borders(xlEdgeTop).Weight = WeightVal
'      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble

'      .Borders(xlEdgeLeft).Weight = WeightVal
'      .Borders(xlEdgeRight).Weight = WeightVal
'      .Borders(xlEdgeTop).Weight = WeightVal
'      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_��d��_��(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'==================================================================================================
Function �r��_��d��_���E(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_��d��_��(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_��d��_�㉺(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDouble
      .Borders(xlEdgeBottom).LineStyle = xlDouble

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'==================================================================================================
Function �r��_�j��_�tL��(Optional setArea As Range, Optional LineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call �r��_�j��_�͂�(setArea, LineColor, WeightVal)
  Call Library.getRGB(LineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    Set setArea = setArea.Offset(1, 1).Resize(setArea.Rows.count - 1, setArea.Columns.count - 1)
    Call �r��_�j��_����(setArea, LineColor, WeightVal)
    Call �r��_�j��_�͂�(setArea, LineColor, WeightVal)
  Else
    setArea.Offset(1, 1).Resize(setArea.Rows.count - 1, setArea.Columns.count - 1).Select
    Call �r��_�j��_����(setArea, LineColor, WeightVal)
    Call �r��_�j��_�͂�(setArea, LineColor, WeightVal)

  End If
End Function



'==================================================================================================
Function �r��_�������폜_��(Optional setArea As Range)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  End If
End Function




'==================================================================================================
Function �r��_�������폜_�c(Optional setArea As Range)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlNone
    End With
  End If

End Function




'**************************************************************************************************
' * �J�������ݒ� / �擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getColumnWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  endColLine = Cells(5, Columns.count).End(xlToLeft).Column

  For colLine = 1 To endColLine
    Cells(1, colLine) = Cells(1, colLine).ColumnWidth
  Next

End Function

'==================================================================================================
Function setColumnWidth()
  Dim colLine As Long, endColLine As Long
  Dim colName As String
  endColLine = Cells(1, Columns.count).End(xlToLeft).Column

  For colLine = 1 To endColLine
    Cells(1, colLine).ColumnWidth = Cells(1, colLine)
  Next

End Function

