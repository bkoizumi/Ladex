VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_VBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'
''�v���C�x�[�g�t�B�[���h�̐錾
'Private WithEvents m_CBE As CommandBarEvents
'
''���������\�b�h
''���̃N���X�̃C���X�^���X������A�����ɌĂ�ŁB
''Java��VB.Net�ł̓R���X�g���N�^�ɏ����������������񂾂��ǁAVBA�ł̓R���X�g���N�^(��Initialize�C�x���g�v���V�[�W��)�Ɉ� �������������Ȃ�����A���������`�ɂ����B
'Public Sub InitializeInstance(ByVal cbc As CommandBarControl)
'    Set m_CBE = Application.VBE.Events.CommandBarEvents(cbc)
'End Sub
'
''�{�^�����N���b�N���ꂽ���̏���
'Private Sub m_CBE_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    Debug.Print "Click" '�C�~�f�B�G�C�g�E�C���h�E��"Click"�Ƃ�����������o��
'End Sub