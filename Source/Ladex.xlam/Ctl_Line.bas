Attribute VB_Name = "Ctl_Line"
Option Explicit

'==================================================================================================
Function �r��_�\_����()
  Call Library.�r��_����_�i�q
End Function


'==================================================================================================
Function �r��_�\_�j��A()
  Call Library.�r��_�\
End Function


'==================================================================================================
Function �r��_�\_�j��B()
  Call Library.�r��_�j��_�i�q
  Call Library.�r��_����_����
  Call Library.�r��_����_�͂�
End Function


'==================================================================================================
Function �r��_�\_�tL��()
  Call init.setting
  Dim startCell As Range, endCell As Range
  
  Set startCell = Selection(1)
  Set endCell = Selection(Selection.count)
  
  Range(startCell.Offset(1, 1), endCell).Select
  Call Library.�r��_�j��_�͂�
  Call Library.�r��_�j��_����
  
  Range(startCell, endCell).Select
  Call Library.�r��_����_�͂�
  
End Function
