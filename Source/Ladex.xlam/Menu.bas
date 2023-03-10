Attribute VB_Name = "Menu"
Option Explicit

'**************************************************************************************************
' * �e�@�\�Ăяo��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �e�@�\�Ăяo��(shortcutName As String)
  
  Const funcName As String = "Menu.�e�@�\�Ăяo��"

  '�����J�n--------------------------------------
  Call init.setting
  If runFlg = False Then
    Call Library.showDebugForm(funcName, , "start")
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg      ", runFlg, "debug")
  Call Library.showDebugForm("shortcutName", shortcutName, "debug")
  '----------------------------------------------

  Select Case shortcutName

    '���C�ɓ���----------------------------------
    Case "Favorite_detail": Call Ctl_Favorite.�ڍו\��
    Case "�ǉ�": Call Ctl_Favorite.�ǉ�

    'Option--------------------------------------
    Case "showOption__": Call Ctl_Option.showOption
    Case "�X�^�C���o��": Call Ctl_Style.�X�^�C���o��
    Case "�X�^�C���捞": Call Ctl_Style.�X�^�C���捞
    Case "showVersion_": Call Ctl_Option.showVersion
    Case "showHelp____": Call Ctl_Option.showHelp
    Case "Addin����___": Call Ctl_Option.Addin����
    Case "������______": Call Ctl_Option.������
    
    '�u�b�N�Ǘ�----------------------------------
    Case "�X�^�C��������___": Call Ctl_Style.�X�^�C��������
    Case "�X�^�C���폜_____": Call Ctl_Style.�X�^�C���폜
    Case "�X�^�C���ݒ�_____": Call Ctl_Style.�X�^�C���ݒ�
    Case "���O��`�폜_____": Call Ctl_Book.���O��`�폜
    Case "B_A1�Z���I��_____": Call Ctl_Book.A1�Z���I��
    Case "B_A1�Z���I��_�ۑ�": Call Ctl_Book.A1�Z���I��_�ۑ�
    Case "����͈͘g�\��___": Call Ctl_Book.����͈͕\��
    Case "����͈͘g��\��_": Call Ctl_Book.����͈͔�\��
    Case "�V�[�g�ꗗ�擾___": Call Ctl_Book.�V�[�g�ꗗ�擾
    Case "�A���V�[�g�ǉ�___": Call Ctl_Book.�A���V�[�g�ǉ�
    Case "�w��{��_________": Call Ctl_Zoom.�w��{��
    Case "�W�����_________": Call Ctl_Book.�W�����
    Case "�V�[�g�Ǘ�_______": Call Ctl_Book.�V�[�g�Ǘ�

    '�V�[�g�Ǘ�----------------------------------
    Case "S_A1�Z���I��______": Call Ctl_Sheet.A1�Z���I��
    Case "S_A1�Z���I��_�ۑ�_": Call Ctl_Sheet.A1�Z���I��_�ۑ�
    Case "���ׂĕ\��________": Call Ctl_Sheet.���ׂĕ\��
    Case "�s�v�f�[�^�폜____": Call Ctl_Sheet.�s�v�f�[�^�폜
    Case "�̍وꊇ�ύX______": Call Ctl_Sheet.�̍وꊇ�ύX
    Case "�t�H���g�ꊇ�ύX__": Call Ctl_Sheet.�t�H���g�ꊇ�ύX

    '���̑��Ǘ�--------------------------------
    Case "�t�@�C�����擾": Call Ctl_File.�t�@�C�����擾
    Case "�t�H���_����____": Call Ctl_File.�t�H���_����
    Case "�摜�\�t��______": Call Ctl_File.�摜�\�t��
    Case "�͂�_�m�F��___": Call Ctl_Stamp.�͂�_�m�F��
    Case "�͂�_���O_____": Call Ctl_Stamp.�͂�_���O
    Case "�͂�_�ψ�_____": Call Ctl_Stamp.�͂�_�ψ�
    Case "QR�R�[�h����____": Call Ctl_shap.QR�R�[�h����
    Case "�p�X���[�h����__": Call Ctl_shap.�p�X���[�h����
    Case "�J�X�^���֐�01__": Call Ctl_�J�X�^��.�J�X�^���֐�01
    Case "�J�X�^���֐�02__": Call Ctl_�J�X�^��.�J�X�^���֐�02
    Case "�J�X�^���֐�03__": Call Ctl_�J�X�^��.�J�X�^���֐�03
    Case "�J�X�^���֐�04__": Call Ctl_�J�X�^��.�J�X�^���֐�04
    Case "�J�X�^���֐�05__": Call Ctl_�J�X�^��.�J�X�^���֐�05


    Case "R1C1�\�L": Call Ctl_Book.R1C1�\�L

    '�Y�[��--------------------------------------
    Case "�S��ʕ\��": Call Ctl_Zoom.�S��ʕ\��



    '�Z������------------------------------------
    Case "�Z����������_��____": Call Ctl_Cells.�Z����������_��
    Case "�Z����������_����__": Call Ctl_Cells.�Z����������_����
    Case "�Z����������_����__": Call Ctl_Cells.�Z����������_����
    Case "�Z���Œ�ݒ�_��____": Call Ctl_Cells.�Z���Œ�ݒ�_��
    Case "�Z���Œ�ݒ�_����__": Call Ctl_Cells.�Z���Œ�ݒ�_����
    Case "�Z���Œ�ݒ�_����__": Call Ctl_Cells.�Z���Œ�ݒ�_����
    Case "�Z���Œ�ݒ�_����15": Call Ctl_Cells.�Z���Œ�ݒ�_����15
    Case "�Z���Œ�ݒ�_����30": Call Ctl_Cells.�Z���Œ�ݒ�_����30

    '�Z���ҏW------------------------------------
    Case "�폜_�O��̃X�y�[�X_________": Call Ctl_Cells.�폜_�O��̃X�y�[�X
    Case "�폜_�S�X�y�[�X_____________": Call Ctl_Cells.�폜_�S�X�y�[�X
    Case "�폜_���s___________________": Call Ctl_Cells.�폜_���s
    Case "�폜_�萔___________________": Call Ctl_Cells.�폜_�萔
    Case "�폜_�R�����g_______________": Call Ctl_Cells.�폜_�R�����g
    Case "�ǉ�_�����ɒ����____________": Call Ctl_Cells.�ǉ�_�����ɒ����_
    Case "�ǉ�_�����ɘA��_____________": Call Ctl_Cells.�ǉ�_�����ɘA��
    Case "�ǉ�_�R�����g_______________": Call Ctl_Cells.�ǉ�_�R�����g
    Case "�㏑_�����ɘA��_____________": Call Ctl_Cells.�㏑_�����ɘA��
    Case "�㏑_�[��___________________": Call Ctl_Cells.�㏑_�[��
    Case "�ϊ�_�S���p_________________": Call Ctl_Cells.�ϊ�_�S�p�˔��p
    Case "�ϊ�_���S�p_________________": Call Ctl_Cells.�ϊ�_���p�ˑS�p
    Case "�ϊ�_�召___________________": Call Ctl_Cells.�ϊ�_�啶���ˏ�����
    Case "�ϊ�_����___________________": Call Ctl_Cells.�ϊ�_�������ˑ啶��
    Case "�ϊ�_URL�G���R�[�h__________": Call Ctl_Cells.�ϊ�_URL�G���R�[�h
    Case "�ϊ�_URL�f�R�[�h____________": Call Ctl_Cells.�ϊ�_URL�f�R�[�h
    Case "�ϊ�_Unicode�G�X�P�[�v______": Call Ctl_Cells.�ϊ�_Unicode�G�X�P�[�v
    Case "�ϊ�_Unicode�A���G�X�P�[�v__": Call Ctl_Cells.�ϊ�_Unicode�A���G�X�P�[�v
    Case "�ϊ�_Base64�G���R�[�h_______": Call Ctl_Cells.�ϊ�_Base64�G���R�[�h
    Case "�ϊ�_Base64�f�R�[�h_________": Call Ctl_Cells.�ϊ�_Base64�f�R�[�h
    Case "�ϊ�_���l���ې���___________": Call Ctl_Cells.�ϊ�_���l�ˊې���
    Case "�ϊ�_�ې����𐔒l___________": Call Ctl_Cells.�ϊ�_�ې����ː��l
    Case "�ݒ�_��������_____________": Call Ctl_Cells.�ݒ�_��������
    Case "�\�t_�s�����ւ�__________": Call Ctl_Cells.�\�t_�s�����ւ�

    '�����ҏW------------------------------------
    Case "�����m�F________________": Call Ctl_Formula.�����m�F
    Case "�����ǉ�_�G���[�h�~_��": Call Ctl_Formula.�G���[�h�~_��
    Case "�����ǉ�_�G���[�h�~_�[��": Call Ctl_Formula.�G���[�h�~_�[��
    Case "�����ǉ�_�[����\��_____": Call Ctl_Formula.�[����\��
    Case "�����ݒ�_�s�ԍ�_________": Call Ctl_Formula.�����ݒ�_�s�ԍ�
    Case "�����ݒ�_�V�[�g��_______": Call Ctl_Formula.�����ݒ�_�V�[�g��

    '���`------------------------------------
    Case "�ړ���T�C�Y�ύX������____": Call Ctl_format.�ړ���T�C�Y�ύX������
    Case "�ړ�����__________________": Call Ctl_format.�ړ�����
    Case "�ړ���T�C�Y�ύX�����Ȃ�__": Call Ctl_format.�ړ���T�C�Y�ύX�����Ȃ�
    Case "�㉺�̗]�����[���ɂ���____": Call Ctl_format.�㉺�]���[��
    Case "���E�̗]�����[���ɂ���____": Call Ctl_format.���E�]���[��
    Case "�����T�C�Y���҂�����ɂ���": Call Ctl_format.�����T�C�Y���҂�����ɂ���
    Case "�Z�����̒����ɔz�u________": Call Ctl_format.�Z�����̒����ɔz�u

    '�摜�ۑ�------------------------------------
    Case "�摜�ۑ�": Call Ctl_Image.�摜�ۑ�

    '�r��[�N���A]--------------------------------
    Case "�r��_�N���A__________": Call Library.�r��_�N���A
    Case "�r��_�N���A_������_��": Call Library.�r��_�N���A_������_��
    Case "�r��_�N���A_������_�c": Call Library.�r��_�N���A_������_�c

    '�r��[�\]------------------------------------
    Case "�r��_�\_����_": Call Ctl_Line.�r��_�\_����
    Case "�r��_�\_�j��A": Call Ctl_Line.�r��_�\_�j��A
    Case "�r��_�\_�j��B": Call Ctl_Line.�r��_�\_�j��B

    '�r��[�j��]----------------------------------
    Case "�r��_�j��_����": Call Library.�r��_�j��_����
    Case "�r��_�j��_����": Call Library.�r��_�j��_����
    Case "�r��_�j��_��__": Call Library.�r��_�j��_��
    Case "�r��_�j��_�E__": Call Library.�r��_�j��_�E
    Case "�r��_�j��_���E": Call Library.�r��_�j��_���E
    Case "�r��_�j��_��__": Call Library.�r��_�j��_��
    Case "�r��_�j��_��__": Call Library.�r��_�j��_��
    Case "�r��_�j��_�㉺": Call Library.�r��_�j��_�㉺
    Case "�r��_�j��_�͂�": Call Library.�r��_�j��_�͂�
    Case "�r��_�j��_�i�q": Call Library.�r��_�j��_�i�q

    '�r��[����]----------------------------------
    Case "�r��_����_����": Call Library.�r��_����_����
    Case "�r��_����_����": Call Library.�r��_����_����
    Case "�r��_����_���E": Call Library.�r��_����_���E
    Case "�r��_����_�㉺": Call Library.�r��_����_�㉺
    Case "�r��_����_�͂�": Call Library.�r��_����_�͂�
    Case "�r��_����_�i�q": Call Library.�r��_����_�i�q

    '�r��[��d��]----------------------------------
    Case "�r��_��d��_��__": Call Library.�r��_��d��_��
    Case "�r��_��d��_�E__": Call Library.�r��_��d��_�E
    Case "�r��_��d��_���E": Call Library.�r��_��d��_���E
    Case "�r��_��d��_��__": Call Library.�r��_��d��_��
    Case "�r��_��d��_��__": Call Library.�r��_��d��_��
    Case "�r��_��d��_�㉺": Call Library.�r��_��d��_�㉺
    Case "�r��_��d��_�͂�": Call Library.�r��_��d��_�͂�

    '�f�[�^����-----------------------------------
    Case "�f�[�^����_�A�ԏ㏑____": Call Ctl_Cells.�㏑_�����ɘA��
    Case "�f�[�^����_�A�Ԓǉ�____": Call Ctl_Cells.�ǉ�_�����ɘA��
    Case "�f�[�^����_�����Œ萔�l": Call Ctl_sampleData.���l_�����Œ�
    Case "�f�[�^����_�͈͎w�萔�l": Call Ctl_sampleData.���l_�͈͎w��
    Case "�f�[�^����_��__________": Call Ctl_sampleData.�f�[�^����_��
    Case "�f�[�^����_��__________": Call Ctl_sampleData.�f�[�^����_��
    Case "�f�[�^����_����________": Call Ctl_sampleData.�f�[�^����_����
    Case "�f�[�^����_���t________": Call Ctl_sampleData.�f�[�^����_���t
    Case "�f�[�^����_����________": Call Ctl_sampleData.�f�[�^����_����
    Case "�f�[�^����_����________": Call Ctl_sampleData.�f�[�^����_����
    Case "�f�[�^����_����________": Call Ctl_sampleData.�f�[�^����_����
    Case "�f�[�^����_�p�^�[���I��": Call Ctl_sampleData.�f�[�^����_�p�^�[���I��
    
    Case Else
      Call Library.showDebugForm("���{�����j���[�Ȃ�", shortcutName, "Error")
      Call Library.showNotice(406, "���{�����j���[�Ȃ��F" & shortcutName, True)
  End Select


  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.resetGlobalVal
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  Exit Function
  '----------------------------------------------

  '�G���[������------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function
