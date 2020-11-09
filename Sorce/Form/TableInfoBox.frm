VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableInfoBox 
   Caption         =   "テーブル情報入力"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8355
   OleObjectBlob   =   "TableInfoBox.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "TableInfoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub OKButton_Click()

  InputTableName = Me.TableName.Value
  InputTableIDa = Me.TableID.Value
  Unload Me
End Sub

Private Sub CancelButton_Click()

  InputTableName = ""
  InputTableID = ""
  
  Unload Me
End Sub
