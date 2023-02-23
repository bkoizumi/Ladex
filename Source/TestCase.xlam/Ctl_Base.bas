Attribute VB_Name = "Ctl_Base"
Option Explicit


'==================================================================================================
Function Chrome�N��(Optional showType As String)

  Dim shellProcID  As Long
  Const funcName As String = "Ctl_Base.Chrome�N��"
  
  'chromeDriver �̍X�V�m�F-------------------------------------------------------------------------
  
  shellProcID = Shell(binPath & "\updateChromeDriver.bat", vbNormalFocus)
  Call Library.chkShellEnd(shellProcID)
  
  Call Library.showDebugForm("targetURL", targetURL, "debug")
  
  With driver
    .AddArgument ("--lang=ja")
    .AddArgument ("--user-data-dir=" & BrowserProfilesDir)
    .AddArgument ("--window-size=1200,1200")
    
    '�g���@�\�𖳌�
    .AddArgument ("--disable-extensions")
    
    
    If setVal("ProxyFlg") = "TRUE" Then
      .AddArgument ("--proxy-server=" & setVal("ProxyURL") & ":" & setVal("ProxyPort"))
      
      If setVal("ProxyID") <> "" Then
        .AddArgument ("--proxy-auth=" & setVal("ProxyID") & ":" & setVal("ProxyPW"))
      End If
    End If
    
    
    'SSL�F�؂̖�����
    .AddArgument ("--ignore-certificate-errors")
    
    
    Select Case showType
      Case "no-img"
        '�摜��\��
        .AddArgument ("--blink-settings=imagesEnabled=false")
      Case "app"
        .AddArgument ("--app=" & Ctl_Base.��{�F��(targetURL))
      
      Case "headless"
        .AddArgument ("--disable-gpu")
        .AddArgument ("--headless")
      
      Case Else
        '�V�[�N���b�g���[�h
        .AddArgument ("--incognito")
    End Select
    
    .start "Chrome"
    
    '��ʈʒu
    '.Window.SetPosition Windows(targetBook.Name).Left, Windows(targetBook.Name).Top
    .Wait 1000
    
    '�y�[�W�̃��[�h�̑҂�����
    .Timeouts.PageLoad = 60000
    
    'javascript���s�����̑҂�����
    .Timeouts.Script = 10000
    
    '�v�f��������܂ł̑҂�����
    .Timeouts.ImplicitWait = 1000
    
    If targetURL <> "" Then
      .Get Ctl_Base.��{�F��(targetURL)
    End If
    
  End With
  
    
End Function

'==================================================================================================
Function Chrome�I��()
  
  Call Library.execDeldir(BrowserProfilesDir)
  driver.Quit
  Set driver = Nothing
  
  Call Library.execMkdir(BrowserProfilesDir)
  
End Function


'==================================================================================================
Function ��{�F��(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.��{�F��"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------

  If InStr(baseURL, setVal("siteMapURL")) <> 0 Then
    Set reg = CreateObject("VBScript.RegExp")
    
    ID = setVal("authName")
    PW = setVal("authPassword")
    
    If setVal("authTypeBasic") = "TRUE" Then
      If (InStr(baseURL, "http://") > 0) Then
        baseURL = Replace(baseURL, "http://", "")
        baseURL = "http://" & ID & ":" & PW & "@" & baseURL
        
      ElseIf (InStr(baseURL, "https://") > 0) Then
        baseURL = Replace(baseURL, "https://", "")
        baseURL = "https://" & ID & ":" & PW & "@" & baseURL
        
      End If
    End If
  End If
  Call Library.showDebugForm("baseURL  ", baseURL, "debug")
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  ��{�F�� = Ctl_Base.�f�t�H���g�y�[�W�ݒ�(baseURL)
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �f�t�H���g�y�[�W�ݒ�(ByVal baseURL As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim ID As String, PW As String
  Dim reg As Object
  
  Const funcName As String = "Ctl_Base.�f�t�H���g�y�[�W�ݒ�"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
'    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
'    Call Library.showDebugForm(funcName, , "start1")
  End If
'  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
 
  If InStr(baseURL, setVal("siteMapURL")) <> 0 Then
    Set reg = CreateObject("VBScript.RegExp")
    
    '/�ŏI����Ă���ꍇ�́AdefaultPage����������
    With reg
      .Pattern = "/$"
      .IgnoreCase = True
      .Global = True
    End With
    
    If reg.Test(baseURL) Then
      baseURL = baseURL & setVal("defaultPage")
      targetURL = targetURL & setVal("defaultPage")
    End If
    Call Library.showDebugForm("baseURL  ", baseURL, "debug")
    
    Set reg = Nothing
  End If
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
'    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
'    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  �f�t�H���g�y�[�W�ݒ� = baseURL
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function




'==================================================================================================
Function ���͕����쐬(ByVal dataVal As Variant)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim inputVal As String
  Dim DataType As Variant, DataLength As Variant, DataFormat As Variant
  
  Const funcName As String = "Ctl_Base.���͕����쐬"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
'�����_���֐��p
'Public Const HalfWidthDigit = "1234567890"
'Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"
'
'Public Const JapaneseCharacters = "�����������������������������������ĂƂȂɂʂ˂̂͂Ђӂւق܂݂ނ߂�������������񂪂����������������������Âłǂ΂тԂׂڂς҂Ղ؂�"
'Public Const JapaneseCharactersCommonUse = "�J�w����щ�⋞���o�m�����X�����������ψ�j�݋��K�n�g�����Ґ̎�󏊒���g�\�������������a�p�ʉ芯�G�����a�ō��Q����������I�T�ŔO�{�@�q��Չ����͋������Ȏ}�ɏq���������Ŕ�񕐉����g���č���@���S�����͓��q���󖇈ˉ���F���������h���������˔t�������ޕ|���b�Ή�����"
'Public Const MachineDependentCharacters = "�@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S�T�U�V�W�X�Y�Z�[�\�]�_�\�]�^�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�z�{"
'

  DataType = dataVal(0)
  DataLength = dataVal(1)
  
  
  
  Select Case DataType
    Case "*"
      
      inputVal = HalfWidthDigit & HalfWidthCharacters & SymbolCharacters & JapaneseCharacters & JapaneseCharactersCommonUse & MachineDependentCharacters
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    Case "day"
      DataFormat = dataVal(2)
      If InStr(DataLength, "y") >= 1 Then
      ElseIf InStr(DataLength, "m") >= 1 Then
        DataLength = Replace(DataLength, "m", "")
        inputVal = DateAdd("m", DataLength, Now())
      Else
        inputVal = DateAdd("d", DataLength, Now())

      End If
      inputVal = Format(inputVal, DataFormat)
    
    
    Case "URL�`��"
      inputVal = LCase(HalfWidthCharacters)
      inputVal = Range("E1") & "/" & Library.makeRandomString(DataLength, inputVal)
    
    Case "�Ђ炪��"
    
      inputVal = JapaneseCharacters
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    Case "�S�p����"
      inputVal = JapaneseCharactersCommonUse
      inputVal = Library.makeRandomString(DataLength, inputVal)
    
    
    Case Else
  End Select



  Call Library.showDebugForm("inputVal", inputVal, "debug")

  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  ���͕����쐬 = inputVal
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
