Attribute VB_Name = "Ctl_Style"
Option Explicit


'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �����ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Const funcName As String = "Ctl_Style.�����ݒ�"

  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Columns("A:A").ColumnWidth = 3
  Columns("B:B").ColumnWidth = 3
    
  '��ƍ���
  Columns("C:G").ColumnWidth = 3
  Columns("H:H").ColumnWidth = 22
  
  '�i����
  With Columns("I:J")
    .ColumnWidth = 6
    .NumberFormatLocal = "0_ ;[��]-0 "
  End With
  
  '�S����
  With Columns("K:L")
    .ColumnWidth = 8
  End With
  
  
  '�\���
  With Columns("M:N")
    .ColumnWidth = 5
    .NumberFormatLocal = "m/d;@"
  End With
  
  '��s/�㑱�^�X�N
  With Columns("O:P")
    .ColumnWidth = 6
    .NumberFormatLocal = "m/d;@"
  End With
  
  
  '���ѓ�
  With Columns("Q:R")
    .ColumnWidth = 6
    .NumberFormatLocal = "m/d;@"
  End With
  
  
  '��ƍH��
  With Columns("S:T")
    .ColumnWidth = 7
    .NumberFormatLocal = "0.0_ ;[��]-0.0 "
  End With
  
  
  '�x���H��
  With Columns("U:U")
    .ColumnWidth = 8
    .NumberFormatLocal = "0.00_ ;[��]-0.00 "
  End With
  
  '���l
  With Columns("V:V")
    .ColumnWidth = 30
  End With

  
  '�J�����_�[����
  With Columns("W:XFD")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .ColumnWidth = 2.5
  End With
  
  Rows(startLine & ":" & Rows.count).RowHeight = 15
  Rows("5:5").RowHeight = 20
  Rows("6:6").RowHeight = 40
  
  With Range("A6:V6")
    .Font.Name = "���C���I"
    .Font.Size = 12
    .Font.Bold = True
    .Font.Strikethrough = False
    .Font.Superscript = False
    .Font.Subscript = False
    .Font.OutlineFont = False
    .Font.Shadow = False
    .Font.Underline = xlUnderlineStyleNone
    .Font.ColorIndex = xlAutomatic
    .Font.TintAndShade = 0
    .Font.ThemeFont = xlThemeFontNone
  End With
  
  Range("W4:XFD4").NumberFormatLocal = "m""��"""
  Range("W5:XFD5").NumberFormatLocal = "d"
  
  
  '�^�X�N�G���A�̏����ݒ�------------------------
  Range("A7:B" & Rows.count).HorizontalAlignment = xlCenter
  Range("C7:H" & Rows.count).HorizontalAlignment = xlGeneral
  
  Range("I7:J" & Rows.count).HorizontalAlignment = xlCenter
  
  '�S����
  With Range("K7:L" & Rows.count)
    .HorizontalAlignment = xlCenter
    .WrapText = False
    .ShrinkToFit = True
  End With
  
  
  Range("M7:R" & Rows.count).HorizontalAlignment = xlCenter
  Range("S7:U" & Rows.count).HorizontalAlignment = xlRight
  
  Range("I7:J" & Rows.count).NumberFormatLocal = "0"
  Range("M7:N" & Rows.count & ",Q4:R" & Rows.count).NumberFormatLocal = "m/d;@"
  'Range("M7:N" & Rows.count & ",Q4:R" & Rows.count).NumberFormatLocal = "mm/dd hh:mm"
  
  Range("Q7:R" & Rows.count).NumberFormatLocal = "m/d;@"


  Range("S7:U" & Rows.count).NumberFormatLocal = "0.00;[��]0.00"
  Range("U7:U" & Rows.count).ShrinkToFit = True
  
  
  
  '���ڃG���A������
  Call Library.�s�v�f�[�^�폜
  endLine = Range("A1").SpecialCells(xlLastCell).Row
  Range("C" & startLine & ":H" & endLine).Merge True
  
  
  
  
  
  
  
  
  
  
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end1")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

