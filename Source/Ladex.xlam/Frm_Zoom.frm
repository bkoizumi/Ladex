VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Zoom 
   Caption         =   "ズーム"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   OleObjectBlob   =   "Frm_Zoom.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Frm_Zoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
Private Sub UserForm_Activate()
 
    Call Ctl_UsrForm.ResizeForm
    
End Sub


'==================================================================================================
Private Sub UserForm_Resize()
  Dim setWidth As Long, setHeight As Long
  
  If Width > 100 Then
    TextBox.Width = Me.Width - 40
    
  End If
  
  If Height > 100 Then
    TextBox.Height = Me.Height - 100
    
    Label1.Top = Me.Height - 75
    Label2.Top = Me.Height - 60
    OK_Button.Top = Me.Height - 75
    CancelButton.Top = Me.Height - 75
  End If
  
End Sub



'==================================================================================================
Private Sub UserForm_Initialize()
  StartUpPosition = 0
  Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
  Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
  
  Caption = "ズーム |  " & thisAppName
  
End Sub

'==================================================================================================
Private Sub CancelButton_Click()
  Unload Me
End Sub

'==================================================================================================
Private Sub OK_Button_Click()
  Call Ctl_Zoom.ZoomOut(TextBox, Frm_Zoom.Label1.Caption)
  Unload Me
End Sub

'==================================================================================================
Private Sub TextBox_Change()
  Call init.setting
  Label2.Caption = "入力文字数：" & Library.getLength(TextBox.Value, "文字数")
End Sub

