VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_smplData 
   Caption         =   "sampleData"
   ClientHeight    =   6012
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525.001
   OleObjectBlob   =   "Frm_smplData.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_smplData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim cmdVal As Variant
  Dim indexCnt As Integer
  
  Call init.setting
  Application.Cursor = xlDefault
  indexCnt = 0
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "データ生成 |  " & thisAppName

  With Frm_smplData
    For Each cmdVal In LadexSh_Config.Range(BK_setVal("Cells_sampleData") & "3:" & BK_setVal("Cells_sampleData") & LadexSh_Config.Cells(Rows.count, 11).End(xlUp).Row)
      ListBox1.AddItem indexCnt & "." & cmdVal
      indexCnt = indexCnt + 1
    Next
  End With
End Sub

'**************************************************************************************************
' * ボタン押下
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub add_Click()
  ListBox2.AddItem ListBox1.Value
  
  If ListBox1.Value Like "*空白*" Then
  Else
    ListBox1.RemoveItem ListBox1.ListIndex
  End If
  
End Sub

'==================================================================================================
Private Sub del_Click()
  If ListBox2.Value Like "*空白*" Then
  Else
    ListBox1.AddItem ListBox2.Value, Split(ListBox2.Value, ".")(0)
  End If
  ListBox2.RemoveItem Me.ListBox2.ListIndex
  
End Sub

'==================================================================================================
Private Sub Cancel_Click()

'  Call Library.setRegistry("UserForm", "mkSmpDtTop", Me.Top)
'  Call Library.setRegistry("UserForm", "mkSmpDtLeft", Me.Left)
  
'  BK_setVal.add "mkSmpDtCancel", "True"
  
  Unload Me
  End
End Sub


'==================================================================================================
Private Sub run_Click()
  
  Call init.setting(True)
'  Call Library.setRegistry("UserForm", "mkSmpDtTop", Me.Top)
'  Call Library.setRegistry("UserForm", "mkSmpDtLeft", Me.Left)
  
  With Me
    Select Case .Caption
      Case "【数値】桁数固定"
        BK_setVal.add "digits", Me.digits1.Text
        BK_setVal.add "maxCount", Me.maxCount1.Text
        BK_setVal.add "addFirst", Me.addFirst.Text
        BK_setVal.add "addEnd", Me.addEnd.Text
        
      Case "【数値】範囲指定"
        BK_setVal.add "maxCount", Me.maxCount2.Text
        
        BK_setVal.add "minVal", Me.minVal2.Text
        BK_setVal.add "maxVal", Me.maxVal2.Text
        BK_setVal.add "addFirst", Me.addFirst.Text
        BK_setVal.add "addEnd", Me.addEnd.Text
      
      Case "【名前】姓", "【名前】名", "【名前】フルネーム"
        BK_setVal.add "maxCount", Me.maxCount3.Text
        
      Case "【日付】日", "【日付】時間", "【日付】日時"
        BK_setVal.add "maxCount", Me.maxCount4.Text
      
        BK_setVal.add "minVal", Me.minVal4.Text
        BK_setVal.add "maxVal", Me.maxVal4.Text
        
      Case "【その他】文字"
        BK_setVal.add "maxCount", Me.maxCount5.Text
      
        BK_setVal.add "strType01", Me.strType01.Value
        BK_setVal.add "strType02", Me.strType02.Value
        BK_setVal.add "strType03", Me.strType03.Value
        BK_setVal.add "strType04", Me.strType04.Value
        BK_setVal.add "strType05", Me.strType05.Value
        BK_setVal.add "strType06", Me.strType06.Value
        BK_setVal.add "strType07", Me.strType07.Value
        
      Case "パターン選択"
        BK_setVal.add "maxCount", Me.maxCount0.Text
        
        Set sampleDataList = Nothing
        Set sampleDataList = CreateObject("Scripting.Dictionary")
        With .Controls("ListBox2")
          For i = 0 To .ListCount - 1
            sampleDataList.add Split(.list(i), ".")(1), Split(.list(i), ".")(1)
          Next
        End With
  
  
      Case Else
    End Select
  End With
    
  

  Unload Me
End Sub

