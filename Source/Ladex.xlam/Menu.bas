Attribute VB_Name = "Menu"
Option Explicit

'**************************************************************************************************
' * �V���[�g�J�b�g�L�[����̌Ăяo���p
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub ladex_Notation_R1C1()
  Call Ctl_Sheet.R1C1�\�L
End Sub

Sub ladex_resetStyle()
  Call Ctl_Style.�X�^�C��������
End Sub

Sub ladex_delStyle()
  Call Ctl_Style.�X�^�C���폜
End Sub

Sub ladex_setStyle()
  Call Ctl_Style.�X�^�C���ݒ�
End Sub

Sub ladex_del_CellNames()
  Call Ctl_Book.���O��`�폜
End Sub

Sub ladex_disp_SVGA12()
  Call Ctl_Window.��ʃT�C�Y�ύX(612, 432)
End Sub

Sub ladex_disp_HD15_6()
  Call Ctl_Window.��ʃT�C�Y�ύX(1920, 1080)
End Sub

Sub ladex_�V�[�g�ꗗ�擾()
  Call Ctl_Book.�V�[�g���X�g�擾
End Sub

Sub ladex_�Z���I��()
  Application.Goto Reference:=Range("A1"), Scroll:=True
End Sub

Sub ladex_�Z���I��_�ۑ�()
  Application.Goto Reference:=Range("A1"), Scroll:=True
  ActiveWorkbook.Save
End Sub

Sub ladex_�S�Z���\��()
  Call Ctl_Sheet.���ׂĕ\��
End Sub

Sub ladex_�Z���ƃV�[�g�I��()
  Call Ctl_Sheet.A1�Z���I��
End Sub

Sub ladex_�Z���ƃV�[�g_�ۑ�()
  Call Ctl_Sheet.A1�Z���I��
  ActiveWorkbook.Save
End Sub

Sub ladex_�W�����()
  Call Ctl_Sheet.�W�����
End Sub

Sub ladex_��ʍő剻()
  Call Ctl_Zoom.Zoom01
End Sub

Sub ladex_�����\���{��()
  Call Ctl_Zoom.defaultZoom
End Sub


Sub ladex_�Z������_��()
  Call Ctl_Cells.�Z��������
End Sub

Sub ladex_�Z������_����()
  Call Ctl_Cells.�Z����������
End Sub

Sub ladex_�Z������_����()
  Call Ctl_Cells.�Z��������
  Call Ctl_Cells.�Z����������
End Sub

Sub ladex_�Z�����擾()
  Call Library.getColumnWidth
End Sub

Sub ladex_�V�[�g�Ǘ�_�t�H�[���\��(Optional control As IRibbonControl)
  Call Ctl_Sheet.�V�[�g�Ǘ�_�t�H�[���\��
End Sub

Sub ladex_���������̃X�y�[�X�폜()
  Call Ctl_Cells.Trim01
End Sub

Sub ladex_�����_�t�^()
  Call Ctl_Cells.�����_�t�^
End Sub

Sub ladex_�A�Ԓǉ�()
  Call Ctl_Cells.�A�Ԓǉ�
End Sub

Sub ladex_�S���p�ϊ�()
  Call Ctl_Cells.�p�����S���p�ϊ�
End Sub

Sub ladex_��������()
  Call Ctl_Cells.���������ݒ�
End Sub

Sub ladex_�R�����g�}��()
  Call init.unsetting
  Call Ctl_Cells.�R�����g�}��
End Sub

Sub ladex_�R�����g�폜()
  Call Ctl_Cells.�R�����g�폜
End Sub

Sub ladex_�s�}��()
  Call Ctl_Cells.�s�}��
End Sub

Sub ladex_��}��()
  Call Ctl_Cells.��}��
End Sub


Sub ladex_�R�����g���`()
  Call Ctl_format.�R�����g���`
End Sub

Sub ladex_�s������ւ��ē\�t��()
  Call Ctl_Cells.�s������ւ��ē\�t��
End Sub

Sub ladex_�����G���[�h�~()
  Call Ctl_Formula.�G���[�h�~
End Sub

Sub ladex_���`_1()
  Call Ctl_format.�ړ���T�C�Y�ύX������
End Sub

Sub ladex_���`_2()
  Call Ctl_format.�ړ�����
End Sub

Sub ladex_���`_3()
  Call Ctl_format.�ړ���T�C�Y�ύX�����Ȃ�
End Sub

Sub ladex_�㉺�̗]���[��()
  Call Ctl_format.�㉺�]���[��
End Sub

Sub ladex_�摜�ۑ�()
  Call Ctl_Image.saveSelectArea2Image
End Sub

Sub ladex_�r��_�N���A()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�N���A
  Call Library.endScript
End Sub

Sub ladex_�r��_�N���A_������_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�������폜_��
  Call Library.endScript
End Sub

Sub ladex_�r��_�N���A_������_�c()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�������폜_�c
  Call Library.endScript
End Sub

Sub ladex_�r��_�\_����()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_�i�q
  Call Library.endScript
End Sub

Sub ladex_�r��_�\_�j��A()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�\
  Call Library.endScript
End Sub

Sub ladex_�r��_�\_�j��B()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_�i�q
  Call Library.�r��_����_����
  Call Library.�r��_����_�͂�
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_����()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_����
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_����()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_����
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_�E()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_�E
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_���E()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_���E
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_�㉺()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_�㉺
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_�͂�()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_�͂�
  Call Library.endScript
End Sub

Sub ladex_�r��_�j��_�i�q()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_�j��_�i�q
  Call Library.endScript
End Sub

Sub ladex_�r��_����_����()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_����
  Call Library.endScript
End Sub

Sub ladex_�r��_����_����()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_����
  Call Library.endScript
End Sub

Sub ladex_�r��_����_���E()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_���E
  Call Library.endScript
End Sub

Sub ladex_�r��_����_�㉺()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_�㉺
  Call Library.endScript
End Sub

Sub ladex_�r��_����_�͂�()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_�͂�
  Call Library.endScript
End Sub

Sub ladex_�r��_����_�i�q()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_����_�i�q
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_���E()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_���E
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_��()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_��
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_�㉺()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_�㉺
  Call Library.endScript
End Sub

Sub ladex_�r��_��d��_�͂�()
  Call init.setting
  Call Library.startScript
  Call Library.�r��_��d��_�͂�
  Call Library.endScript
End Sub

Sub ladex_�A�Ԑݒ�()
  Call Ctl_Cells.�A�Ԑݒ�
End Sub

Sub ladex_�A�Ԑ���()
  Call Ctl_Cells.�A�Ԓǉ�
End Sub

Sub ladex_�����Œ萔�l()
  Call Ctl_sampleData.���l_�����Œ�(Selection.CountLarge)
End Sub

Sub ladex_�͈͎w�萔�l()
  Call Ctl_sampleData.���l_�͈�
End Sub

Sub ladex_�f�[�^����_��()
  Call Ctl_sampleData.���O_��(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_��()
  Call Ctl_sampleData.���O_��(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_����()
  Call Ctl_sampleData.���O_�t���l�[��(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_���t()
  Call Ctl_sampleData.���t_��(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_����()
  Call Ctl_sampleData.���t_����(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_����()
  Call Ctl_sampleData.����(Selection.CountLarge)
End Sub

Sub ladex_�f�[�^����_����()
  Call Ctl_sampleData.���̑�_����
End Sub

Sub ladex_�J�X�^��01()
  Call Ctl_�J�X�^��.�J�X�^��01
End Sub
Sub ladex_�J�X�^��02()
  Call Ctl_�J�X�^��.�J�X�^��02
End Sub
Sub ladex_�J�X�^��03()
  Call Ctl_�J�X�^��.�J�X�^��03
End Sub
Sub ladex_�J�X�^��04()
  Call Ctl_�J�X�^��.�J�X�^��04
End Sub
Sub ladex_�J�X�^��05()
  Call Ctl_�J�X�^��.�J�X�^��05
End Sub


