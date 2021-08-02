VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_smplData 
   Caption         =   "sampleData"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525
   OleObjectBlob   =   "Frm_smplData.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_smplData"
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
  Dim cmdVal As Variant
  Dim indexCnt As Integer
  
  Call init.setting
  Application.Cursor = xlDefault
  indexCnt = 0
  
  With Frm_smplData
    For Each cmdVal In BK_sheetSetting.Range(BK_setVal("Cells_sampleData") & "3:" & BK_setVal("Cells_sampleData") & "22")
      ListBox1.AddItem indexCnt & "." & cmdVal
      indexCnt = indexCnt + 1
    Next
  End With
End Sub





'**************************************************************************************************
' * �{�^������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub add_Click()
  ListBox2.AddItem Me.ListBox1.Value
  ListBox1.RemoveItem Me.ListBox1.ListIndex
  
End Sub

'==================================================================================================
Private Sub del_Click()
  ListBox1.AddItem Me.ListBox2.Value, Split(Me.ListBox2.Value, ".")(0)
  ListBox2.RemoveItem Me.ListBox2.ListIndex
  
End Sub

'==================================================================================================
Private Sub Cancel_Click()

  Call Library.setRegistry("UserForm", "mkSmpDtTop", Me.Top)
  Call Library.setRegistry("UserForm", "mkSmpDtLeft", Me.Left)
  
'  BK_setVal.add "mkSmpDtCancel", "True"
  
  Unload Me
  End
End Sub


'==================================================================================================
Private Sub run_Click()
  
  Call init.setting(True)
  Call Library.setRegistry("UserForm", "mkSmpDtTop", Me.Top)
  Call Library.setRegistry("UserForm", "mkSmpDtLeft", Me.Left)
  
  With Me
    Select Case .Caption
      Case "�y���l�z�����Œ�"
        BK_setVal.add "digits", Me.digits1.Text
        BK_setVal.add "maxCount", Me.maxCount1.Text
        
      Case "�y���l�z�͈͎w��"
        BK_setVal.add "maxCount", Me.maxCount2.Text
        
        BK_setVal.add "minVal", Me.minVal2.Text
        BK_setVal.add "maxVal", Me.maxVal2.Text
      
      Case "�y���O�z��", "�y���O�z��", "�y���O�z�t���l�[��"
        BK_setVal.add "maxCount", Me.maxCount3.Text
        
      Case "�y���t�z��", "�y���t�z����", "�y���t�z����"
        BK_setVal.add "maxCount", Me.maxCount4.Text
      
        BK_setVal.add "minVal", Me.minVal4.Text
        BK_setVal.add "maxVal", Me.maxVal4.Text
        
      Case "�p�^�[���I��"
        BK_setVal.add "maxCount", Me.maxCount0.Text
        
        Set sampleDataList = Nothing
        Set sampleDataList = CreateObject("Scripting.Dictionary")
        With .Controls("ListBox2")
          For i = 0 To .ListCount - 1
            sampleDataList.add "list_" & i, .list(i)
          Next
        End With
  
  
      Case Else
    End Select
  End With
    
  

  Unload Me
End Sub

