VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ctl_Event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'https://excel-ubara.com/vba_class/VBA_CLASS_02.html

Option Explicit

'�Ώۂ�Excel�A�v���P�[�V����
Private WithEvents xlApp As Excel.Application
Attribute xlApp.VB_VarHelpID = -1

'�Ώۃu�b�N�E�V�[�g�F1�����z��̗v�f�ɍ\���̂�����
Private Type typeBookSheets
    Book As String '�u�b�N
    Sheets() As String '�V�[�g1,�V�[�g2,�c
End Type
Private pBookSheets() As typeBookSheets
Private pCntBS As Integer

'�C�x���g��~�i�ꎞ�I�ɒ�~������j
Private pStopEvent As Boolean


'==================================================================================================
' ���J�v���p�e�B
'==================================================================================================
'Excel�A�v���P�[�V�����̐ݒ�E�擾
Public Property Set ExcelApplication(argApp As Excel.Application)
    Set xlApp = argApp
End Property

Public Property Get ExcelApplication() As Excel.Application
    Set ExcelApplication = xlApp
End Property

'�C�x���g��~�i�ꎞ�I�ɒ�~������j�̎擾�ݒ�
Public Property Let StopEvent(argStop As Boolean)
    pStopEvent = argStop
End Property

Public Property Get StopEvent() As Boolean
  StopEvent = pStopEvent
End Property

'==================================================================================================
' ���J���\�b�h
'==================================================================================================
'�Ώۃu�b�N�E�V�[�g�̏�����
Public Sub InitializeBookSheets()
    pCntBS = 0
    Erase pBookSheets
End Sub


'==================================================================================================
'�Ώۃu�b�N�E�V�[�g�̐ݒ�
Public Function AddBookSheet(ByRef argAry) As Boolean
    Dim i1 As Long
    Dim i2 As Long
    Dim ix As Integer
    Dim myArray() As String

    '�z��̎������ɂ�蕪��
    Select Case GetDimension(argAry)

        Case 1 '1�����z��
            ReDim Preserve pBookSheets(pCntBS)
            pBookSheets(pCntBS).Book = argAry(LBound(argAry))
            ix = 0
            ReDim myArray(ix)
            For i1 = LBound(argAry) + 1 To UBound(argAry)
                If argAry(i1) <> "" Then
                    ReDim Preserve myArray(ix)
                    myArray(ix) = argAry(i1)
                    ix = ix + 1
                End If
            Next
            pBookSheets(pCntBS).Sheets = myArray
            Erase myArray
            pCntBS = pCntBS + 1

        Case 2 '2�����z��
            For i1 = LBound(argAry, 1) To UBound(argAry, 1)
                ReDim Preserve pBookSheets(pCntBS)
                pBookSheets(pCntBS).Book = argAry(i1, LBound(argAry, 2))
                ix = 0
                ReDim myArray(ix)
                For i2 = LBound(argAry, 2) + 1 To UBound(argAry, 2)
                    If argAry(i1, i2) <> "" Then
                        ReDim Preserve myArray(ix)
                        myArray(ix) = argAry(i1, i2)
                        ix = ix + 1
                    End If
                Next
                pBookSheets(pCntBS).Sheets = myArray
                Erase myArray
                pCntBS = pCntBS + 1
            Next

        Case Else '1,2�����ȊO�̓G���[�Ŗ߂�
            AddBookSheet = False
            Exit Function
    End Select

    AddBookSheet = True
End Function


'==================================================================================================
'�Ώۃu�b�N�E�V�[�g�̑S�擾�F2�����z��Ŗ߂�
Public Function GetBookSheets() As Variant
    Dim i1 As Integer, i2 As Integer
    Dim intMax As Integer
    Dim myArray() As String

    '�z����̍ő�V�[�g�����擾
    intMax = 0
    For i1 = 0 To UBound(pBookSheets)
        If UBound(pBookSheets(i1).Sheets) > intMax Then
            intMax = UBound(pBookSheets(i1).Sheets)
        End If
    Next

    '�\���̂�1�����z����A�\���̂��o������2�����z���
    ReDim myArray(UBound(pBookSheets), intMax + 1)
    For i1 = 0 To UBound(pBookSheets)
        myArray(i1, 0) = pBookSheets(i1).Book
        For i2 = LBound(pBookSheets(i1).Sheets) To UBound(pBookSheets(i1).Sheets)
            myArray(i1, i2 + 1) = pBookSheets(i1).Sheets(i2)
        Next
    Next
    GetBookSheets = myArray
End Function

'==================================================================================================
' �N���X�̃��\�b�h
'==================================================================================================

'�N���X�̏�����
Private Sub Class_Initialize()
    Set Me.ExcelApplication = Application
    Me.StopEvent = False
    Me.InitializeBookSheets
End Sub

'�N���X�̏I������
Private Sub Class_Terminate()
    Set xlApp = Nothing
End Sub

'==================================================================================================
' Application�̃u�b�N�C�x���g
'==================================================================================================


'==================================================================================================
'�u�b�N���J���ꂽ�Ƃ��ɔ���
Private Sub xlApp_WorkbookOpen(ByVal wb As Workbook)
  Const funcName As String = "Ctl_Event.xlApp_WorkbookOpen"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  runFlg = True
'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.startScript
'  Call Main.setShortcutKey
'  Call Library.endScript
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub


'==================================================================================================
'�u�b�N�������钼�O�ɔ���
Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_WorkbookBeforeClose"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
'�V�����u�b�N���쐬���ꂽ�Ƃ��ɔ���
Private Sub xlApp_NewWorkbook(ByVal wb As Workbook)
  Const funcName As String = "Ctl_Event.xlApp_NewWorkbook"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckBook(wb) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
' Application�̃V�[�g�C�x���g
'==================================================================================================


'==================================================================================================
'�V�[�g���A�N�e�B�u�ɂȂ����Ƃ��ɔ���
Private Sub xlApp_SheetActivate(ByVal sh As Object)
  Const funcName As String = "Ctl_Event.xlApp_SheetActivate"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm(funcName, , "end")
  
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
'���[�N�V�[�g�̃Z�����ύX���ꂽ�Ƃ��ɔ���
Private Sub xlApp_SheetChange(ByVal sh As Object, ByVal Target As Range)
  Const funcName As String = "Ctl_Event.xlApp_SheetChange"
    
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub

'  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
'  Call Library.showDebugForm(funcName, , "end")

  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
Private Sub xlApp_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
  Const funcName As String = "Ctl_Event.xlApp_SheetSelectionChange"
  
  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub
  
  Call init.setting
'  Call Library.showDebugForm(funcName, , "start")
  
  '�n�C���C�g�@�\--------------------------------
  If BKh_rbPressed = True Then
    runFlg = True
    Call Ctl_HighLight.showStart(Target)
  End If
  
  '�����m�F�@�\----------------------------------
  If BKcf_rbPressed = True Then
    runFlg = True
    Call Ctl_Formula.�����m�F
  End If
  
'  Call Library.showDebugForm(funcName, , "end")
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub


'==================================================================================================
'�Z�����_�u���N���b�N���ꂽ�Ƃ��ɔ���
Private Sub xlApp_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
  Const funcName As String = "Ctl_Event.xlApp_SheetBeforeDoubleClick"

  On Error GoTo catchError
  If Me.StopEvent Then Exit Sub
  If Not CheckSheet(sh) Then Exit Sub

  Call init.setting
  Call Library.showDebugForm(funcName, , "start")
  
  '�Y�[���@�\------------------------------------
  If BKz_rbPressed = True Then
    Cancel = True
    runFlg = True
    Call Ctl_Zoom.ZoomIn(CStr(Target.Address(False, False)))
  End If
  Call Library.showDebugForm(funcName, , "end")
  Exit Sub
  
'�G���[������------------------------------------
catchError:
  Debug.Print Format(Now(), "yyyy-mm-dd hh:nn:ss") & "  [ERROR]" & funcName & "  " & Err.Description
End Sub

'==================================================================================================
' ����J���\�b�h
'==================================================================================================


'==================================================================================================
'�Ώۃu�b�N�̔���
Private Function CheckBook(ByVal wb As Workbook) As Boolean
  '�w��Ȃ��͂��ׂđΏۂƂ���
  If pCntBS = 0 Then
    CheckBook = True
    Exit Function
  End If
  
  '�u�b�N�E�V�[�g�w��̔z��Ƀu�b�N�������邩����
  Dim i1 As Integer
  CheckBook = False
  For i1 = 0 To UBound(pBookSheets)
    '�u�b�N�̖��̔���
    If wb.Name Like pBookSheets(i1).Book Then
      '�V�[�g�w�肪�Ȃ����̂݃u�b�N�̎w��Ɣ���
      If pBookSheets(i1).Sheets(0) = "" Then
        CheckBook = True
        Exit Function
      End If
    End If
  Next
End Function

'==================================================================================================
'�ΏۃV�[�g�̔���
Private Function CheckSheet(ByVal ws As Worksheet) As Boolean
    
  '�w��Ȃ��͂��ׂđΏۂƂ���
  If pCntBS = 0 Then
    CheckSheet = True
    Exit Function
  End If

  '�u�b�N�E�V�[�g�w��̔z��ɃV�[�g�������邩����
  Dim i1 As Integer, i2 As Integer
  Dim wb As Workbook
  CheckSheet = False
  Set wb = ws.Parent
  For i1 = 0 To UBound(pBookSheets)
    '�u�b�N�̖��̔���
    If wb.Name Like pBookSheets(i1).Book Then
      '�V�[�g�̖��̔���
      For i2 = LBound(pBookSheets(i1).Sheets) To UBound(pBookSheets(i1).Sheets)
        If ws.Name Like pBookSheets(i1).Sheets(i2) Then
          CheckSheet = True
          Exit Function
        End If
      Next
    End If
  Next
End Function


'==================================================================================================
'�z��̎������擾�ƃG���[�l����F1,2�ȊO�̖߂�l�̓G���[
Private Function GetDimension(ByVal argAry) As Integer
  On Error Resume Next
  
  '�z��̎������擾
  Dim intDim As Integer
  Dim tmpDim As Integer
  intDim = 0
  Do While Err.Number = 0
      intDim = intDim + 1
      '�G���[�ɂȂ������_�Ŕz�񂪏I���
      tmpDim = UBound(argAry, intDim)
  Loop
  Err.Clear
  intDim = intDim - 1
  
  '�z��̏ꍇ�A�G���[�l�f�[�^�̔���
  Dim vTemp
  Select Case intDim
    Case 1, 2
      For Each vTemp In argAry
        If IsError(vTemp) Then
          GetDimension = 0
          Exit Function
        End If
      Next
  End Select
  
  GetDimension = intDim
  On Error GoTo 0
End Function