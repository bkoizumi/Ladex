VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_HighLight 
   Caption         =   "HighLight"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   OleObjectBlob   =   "Frm_HighLight.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Frm_HighLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 And Win64 Then
    Public hwnd As LongPtr
#Else
    Public hwnd As Long
#End If

'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub UserForm_Initialize()
  
  hwnd = FindWindow("ThunderDFrame", Me.Caption)
  
  If hwnd <> 0& Then
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or &H20
    
    'フレーム無
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    
    'キャプションなし
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION
    
    '半透明化
    SetLayeredWindowAttributes hwnd, 0, 120, LWA_ALPHA
  End If
  
  

 
End Sub



'==================================================================================================
Sub close_b()

Unload Me
End Sub

