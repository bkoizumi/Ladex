Attribute VB_Name = "Ctl_Cells"
'**************************************************************************************************
' * �Z���ҏW
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' * �����񑀍�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Trim01()
  
  Call init.setting
  ActiveCell = Trim(ActiveCell.Text)
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:

End Function

'==================================================================================================
Function �����_�t�^()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^�E"
    .IgnoreCase = False
    .Global = True
  End With
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Cells.�����_�t�^"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  For Each slctCells In Selection
    'Call Library.showDebugForm("�I���Z���l�F" & Reg.Replace(slctCells.Value, ""))
    slctCells.Value = "�E" & Reg.Replace(slctCells.Value, "")
    
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
Function �A�Ԑݒ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Cells.�A�Ԑݒ�"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
    'slctCells.NumberFormatLocal = "@"
    slctCells.Value = line
    line = line + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function



'==================================================================================================
Function �A�Ԓǉ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  Dim Reg As Object
  
  Set Reg = CreateObject("VBScript.RegExp")
  With Reg
    .Pattern = "^[0-9]+�D"
    .IgnoreCase = False
    .Global = True
  End With
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Cells.�A�Ԓǉ�"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
'    Call Library.showDebugForm("�I���Z���l�F" & Reg.Replace(slctCells.Value, ""))
    slctCells.Value = line & "�D" & Reg.Replace(slctCells.Value, "")
    line = line + 1
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function



'==================================================================================================
Function �p�����S���p�ϊ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Cells.�p�����S���p�ϊ�"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  For Each slctCells In Selection
    slctCells.Value = Library.convHan2Zen(slctCells.Value)
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
Function ���������ݒ�()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  funcName = "Ctl_Cells.���������ݒ�"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  For Each slctCells In Selection
    If slctCells.Font.Strikethrough = True Then
      slctCells.Font.Strikethrough = False
    Else
      slctCells.Font.Strikethrough = True
    End If
    DoEvents
  Next

  '�����I��--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function

'==================================================================================================
Function �R�����g�}��()
  Dim commentVal As String
  Const funcName As String = "Ctl_Cells.�R�����g�}��"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  commentVal = ""
  If TypeName(ActiveCell.Comment) = "Comment" Then
    commentVal = ActiveCell.Comment.Text
  End If
  With Frm_InsComment
    .TextBox = commentVal
    .Label1.Caption = "�I���Z���F" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    .Show
  End With
  
  '�����I��--------------------------------------
  Call Library.showDebugForm("  ", , "end")
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function

'==================================================================================================
Function �R�����g�폜()
  Const funcName As String = "Ctl_Cells.�R�����g�폜"

  '�����J�n--------------------------------------
  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("" & funcName, , "function")
  '----------------------------------------------
  
  If TypeName(ActiveCell.Comment) = "Comment" Then
    ActiveCell.ClearComments
  End If
  
  '�����I��--------------------------------------
  Call Library.showDebugForm("  ", , "end")
  '----------------------------------------------

  Exit Function
'�G���[������------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & " [" & Err.Number & "]" & Err.Description & ">", True)
End Function


