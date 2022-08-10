VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Info 
   Caption         =   "情報"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "Frm_Info.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Info"
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
  Dim objSheet As Object
  Dim cBox As CommandBarComboBox
  
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  Caption = "情報 |  " & thisAppName
  
  Frm_Info.copySheet.AddItem "≪新規シート≫"
  For Each objSheet In ActiveWorkbook.Sheets
    If objSheet.Visible = True Then
      Frm_Info.copySheet.AddItem objSheet.Name
    End If
  Next
  Frm_Info.copySheet.ListIndex = 0
  Set cBox = Nothing
  
End Sub


'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

'==================================================================================================
'OKボタン押下処理
Private Sub OKButton_Click()
  
  Set FrmVal = Nothing
  Set FrmVal = CreateObject("Scripting.Dictionary")
  
  FrmVal.add "SheetList", TextBox.Value
  FrmVal.add "copySheet", copySheet.Value
  
  Unload Me
End Sub


