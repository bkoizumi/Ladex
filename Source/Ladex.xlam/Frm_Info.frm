VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Info 
   Caption         =   "���"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "Frm_Info.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim objSheet As Object
  Dim cBox As CommandBarComboBox
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "��� |  " & thisAppName
  
  Frm_Info.copySheet.AddItem "��V�K�V�[�g��"
  For Each objSheet In ActiveWorkbook.Sheets
    If objSheet.Visible = True Then
      Frm_Info.copySheet.AddItem objSheet.Name
    End If
  Next
  Frm_Info.copySheet.ListIndex = 0
  Set cBox = Nothing
  
End Sub


'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
'OK�{�^����������
Private Sub OKButton_Click()
  
  Set FrmVal = Nothing
  Set FrmVal = CreateObject("Scripting.Dictionary")
  
  FrmVal.add "SheetList", TextBox.Value
  FrmVal.add "copySheet", copySheet.Value
  
  Unload Me
End Sub


