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
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
'// 32bit��
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If


'**************************************************************************************************
' * �t�H�[���T�C�Y�ύX
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function FormResize()
  Dim hWnd  As Long, style As Long
  
  '�E�C���h�E�n���h���擾
  hWnd = GetActiveWindow()
  
  '�E�C���h�E�̃X�^�C�����擾
  style = GetWindowLong(hWnd, GWL_STYLE)
  
  '�E�C���h�E�̃X�^�C���ɃE�C���h�E�T�C�Y�ρ{�ŏ��{�^���{�ő�{�^����ǉ�
  style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
  
  '�E�C���h�E�̃X�^�C�����Đݒ�
  Call SetWindowLong(hWnd, GWL_STYLE, style)
End Function


'**************************************************************************************************
' * ���[�U�[��`�֐�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeUsrFunction()
  Const funcName As String = "Ctl_UsrFunction.InitializeUsrFunction"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
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
    ArgumentDescriptions:=Array("�`�F�b�N������t���w��", "��N�T�𐔒l�Ŏw��", "�j���𐔒l�Ŏw��" & vbNewLine & _
                                "1�F���@2�F�΁@3�F��" & vbNewLine & _
                                "4�F�؁@5�F���@6�F�y�@7�F��")

  Application.MacroOptions _
    Macro:="Textjoin", _
    Description:="������A��", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("��؂蕶��", "�󗓎�����[True�F��������/False�F�������Ȃ�]", "������1,������2, ...")

  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function Textjoin(Delim As String, ParamArray par())
Attribute Textjoin.VB_Description = "������A��"
Attribute Textjoin.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim i As Integer
  Dim tR As Range

'  Application.Volatile

  Textjoin = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Then
          Textjoin = Textjoin & Delim & tR.Value2
        End If
      Next
    Else
      If (par(i) <> "" And par(i) <> "<>") Then
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


'==================================================================================================
'Function mkQRcode(ByVal codeVal As String, Optional ByVal QRSize As Long = 140) As String
'  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  Dim slctCells As Range
'
'  Dim chartAPIURL As String
'  Dim QRCodeImgName As String
'
'  Const funcName As String = "Ctl_Shap.QR�R�[�h����"
'  Const chartAPI = "https://chart.googleapis.com/chart?cht=qr&chld=l|1&"
'
'
'  '�����J�n--------------------------------------
''  On Error GoTo catchError
'  Call init.setting
'  Call Library.startScript
'  '----------------------------------------------
'
'  Set slctCells = Application.Caller
'  QRCodeImgName = "QRCode_" & Application.Caller.Address(False, False)
'
'  Call Library.showDebugForm("slctCells", slctCells.Address, "debug")
'  Call Library.showDebugForm("codeVal", codeVal, "debug")
'  Call Library.showDebugForm("QRCodeImgName", QRCodeImgName, "debug")
'
'
'  If Library.chkShapeName(QRCodeImgName) Then
'    ActiveSheet.Shapes.Range(Array(QRCodeImgName)).Select
'    Selection.delete
'  End If
'
'
'  chartAPIURL = chartAPI & "chs=" & QRSize & "x" & QRSize
'  chartAPIURL = chartAPIURL & "&chl=" & Library.convURLEncode(codeVal)
'  Call Library.showDebugForm("chartAPIURL", chartAPIURL, "debug")
'
'  With ActiveSheet.Pictures.Insert(chartAPIURL)
'    .ShapeRange.Top = slctCells.Top + (slctCells.Height - .ShapeRange.Height) / 2
'    .ShapeRange.Left = slctCells.Left + (slctCells.Width - .ShapeRange.Width) / 2
'
'    .Placement = xlMove
'
'    'QR�R�[�h�̖��O�ݒ�
'    .ShapeRange.Name = QRCodeImgName
'    .Name = QRCodeImgName
'  End With
'
'  mkQRcode = ""
''  ActiveSheet.Select
''  slctCells.Select
'  Set slctCells = Nothing
'
'  '�����I��--------------------------------------
'Lbl_endFunction:
'  Call Library.endScript
'  Call Library.showDebugForm(funcName, , "end")
'  Call init.unsetting
'  '----------------------------------------------
'  Exit Function
'
''�G���[������------------------------------------
'catchError:
'  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
'End Function


