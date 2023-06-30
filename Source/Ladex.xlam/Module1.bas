Attribute VB_Name = "Module1"
Sub test()

  For Each objStyle In ActiveWorkbook.Styles
    Debug.Print objStyle.Name
  Next
  
  
End Sub
