Attribute VB_Name = "Ctl_Line"
Option Explicit

'==================================================================================================
Function rü_\_Àü()
  Call Library.rü_Àü_iq
End Function


'==================================================================================================
Function rü_\_jüA()
  Call Library.rü_\
End Function


'==================================================================================================
Function rü_\_jüB()
  Call Library.rü_jü_iq
  Call Library.rü_Àü_½
  Call Library.rü_Àü_ÍÝ
End Function


'==================================================================================================
Function rü_\_tL()
  Call init.setting
  Dim startCell As Range, endCell As Range
  
  Set startCell = Selection(1)
  Set endCell = Selection(Selection.count)
  
  Range(startCell.Offset(1, 1), endCell).Select
  Call Library.rü_jü_ÍÝ
  Call Library.rü_jü_½
  
  Range(startCell, endCell).Select
  Call Library.rü_Àü_ÍÝ
  
End Function
