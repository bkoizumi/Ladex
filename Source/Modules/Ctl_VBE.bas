Attribute VB_Name = "Ctl_VBE"
Option Explicit

Private Cls_VBE As Cls_VBE

Private cbc As CommandBarControl

Sub addButton()
  
  Set cbc = Application.VBE.CommandBars("Ladex").Controls.add(Type:=msoControlButton, ID:=1, Before:=1)
  Set Cls_VBE = New Cls_VBE
  Call Cls_VBE.InitializeInstance(m_CBB)
'  CBC.FaceId = 444
  
  
End Sub


Sub deleteButton()
    On Error Resume Next
    
    Call cbc.delete
End Sub

Sub test()
    Debug.Print "test"
    MsgBox "test"
    
End Sub



Sub �S�ẴR�[�h�E�C���h�E�����()
    Dim c As CodePane
    For Each c In Application.VBE.CodePanes
        c.Window.Close
    Next
End Sub


Sub �C�~�f�B�G�C�g�E�B���h�E���N���A()
    Debug.Print String(200, vbCrLf)
End Sub
