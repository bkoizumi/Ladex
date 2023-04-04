Attribute VB_Name = "Ctl_Line"
Option Explicit

'==================================================================================================
Function Œrü_•\_Àü()
  Call Library.Œrü_Àü_Šiq
End Function


'==================================================================================================
Function Œrü_•\_”jüA()
  Call Library.Œrü_•\
End Function


'==================================================================================================
Function Œrü_•\_”jüB()
  Call Library.Œrü_”jü_Šiq
  Call Library.Œrü_Àü_…•½
  Call Library.Œrü_Àü_ˆÍ‚İ
End Function


'==================================================================================================
Function Œrü_•\_‹tLš()
  Call init.setting
  Dim startCell As Range, endCell As Range
  
  Set startCell = Selection(1)
  Set endCell = Selection(Selection.count)
  
  Range(startCell.Offset(1, 1), endCell).Select
  Call Library.Œrü_”jü_ˆÍ‚İ
  Call Library.Œrü_”jü_…•½
  
  Range(startCell, endCell).Select
  Call Library.Œrü_Àü_ˆÍ‚İ
  
End Function
