Attribute VB_Name = "Ctl_UsrFunction"
Option Explicit

'// Win32API�p�萔
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
'// Win32API�Q�Ɛ錾
'// 64bit��
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
'// 32bit��
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If


'**************************************************************************************************
' * �t�H�[���T�C�Y�ύX
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function FormResize()
    Dim hwnd                    As Long         '// �E�C���h�E�n���h��
    Dim style                   As Long         '// �E�C���h�E�X�^�C��
 
    '// �E�C���h�E�n���h���擾
    hwnd = GetActiveWindow()
    
    '// �E�C���h�E�̃X�^�C�����擾
    style = GetWindowLong(hwnd, GWL_STYLE)
    
    '// �E�C���h�E�̃X�^�C���ɃE�C���h�E�T�C�Y�ρ{�ŏ��{�^���{�ő�{�^����ǉ�
    style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
 
    '// �E�C���h�E�̃X�^�C�����Đݒ�
    Call SetWindowLong(hwnd, GWL_STYLE, style)
End Function


'**************************************************************************************************
' * ���[�U�[��`�֐�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeUsrFunction()

  Application.MacroOptions _
    Macro:="chkWorkDay", _
    Description:="��N�c�Ɠ����`�F�b�N���ATrue/False��Ԃ�", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("�`�F�b�N������t���w��", "��N�c�Ɠ��𐔒l�Ŏw��")

  Application.MacroOptions _
    Macro:="getWorkDay", _
    Description:="��N�c�Ɠ����V���A���l�ŕԂ�", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("�`�F�b�N����N�𐔒l�Ŏw��", "�`�F�b�N���錎�𐔒l�Ŏw��", "��N�c�Ɠ��𐔒l�Ŏw��")

  Application.MacroOptions _
    Macro:="chkWeekNum", _
    Description:="��N�TX�j���̓��t���`�F�b�N���ATrue/False��Ԃ�", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("�`�F�b�N������t���w��", "��N�T�𐔒l�Ŏw��", "X�j���𐔒l�Ŏw��")

  Application.MacroOptions _
    Macro:="Textjoin", _
    Description:="������A��", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("��؂蕶��", "�󗓎�����[True�F��������/False�F�������Ȃ�]", "������1,������2, ...")


End Function



'==================================================================================================
Function Textjoin(Delim, Ignore As Boolean, ParamArray par())
Attribute Textjoin.VB_Description = "������A��"
Attribute Textjoin.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim i As Integer
  Dim tR As Range

'  Application.Volatile

  Textjoin = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Or Ignore = False Then
          Textjoin = Textjoin & Delim & tR.Value2
        End If
      Next
    Else
      If (par(i) <> "" And par(i) <> "<>") Or Ignore = False Then
        Textjoin = Textjoin & Delim & par(i)
      End If
    End If
  Next

  Textjoin = Mid(Textjoin, Len(Delim) + 1)

End Function




'==================================================================================================
Function chkWorkDay(ByVal checkDate As Date, ByVal bsnDay As Long) As Boolean
Attribute chkWorkDay.VB_Description = "��N�c�Ɠ����`�F�b�N���ATrue/False��Ԃ�"
Attribute chkWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
  
'  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(Year(checkDate), Month(checkDate), 1)
  getDay = Application.WorksheetFunction.WorkDay(firstDay, bsnDay, arryHollyday)
  
  If checkDate = getDay Then
    chkWorkDay = True
  Else
    chkWorkDay = False
  End If

End Function

'==================================================================================================
Function chkWeekNum(ByVal checkDate As Date, ByVal checkWeekday As Long, ByVal weekNum As Long) As Boolean
Attribute chkWeekNum.VB_Description = "��N�TX�j���̓��t���`�F�b�N���ATrue/False��Ԃ�"
Attribute chkWeekNum.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Long, diff As Long
  
'  Application.Volatile
  
  firstDay = Weekday(DateSerial(Year(checkDate), Month(checkDate), 1))
  diff = (checkWeekday + 7 - firstDay) Mod 7
  getDay = DateSerial(Year(checkDate), Month(checkDate), 1 + diff + 7 * (weekNum - 1))
  
  If checkDate = getDay Then
    chkWeekNum = True
  Else
    chkWeekNum = False
  End If
  
End Function

'==================================================================================================
Function getWorkDay(ByVal cYear As Long, ByVal cMonth As Long, ByVal bsnDay As Long) As Date
Attribute getWorkDay.VB_Description = "��N�c�Ɠ����V���A���l�ŕԂ�"
Attribute getWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
'  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(cYear, cMonth, 1)
  getWorkDay = Application.WorksheetFunction.WorkDay(firstDay - 1, bsnDay, arryHollyday)
  
End Function

