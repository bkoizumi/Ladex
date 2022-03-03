Attribute VB_Name = "Ctl_format"
Option Explicit

'**************************************************************************************************
' * �R�����g���`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �R�����g���`()
  Const funcName As String = "Ctl_Format.�R�����g���`"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  If TypeName(ActiveCell) = "Range" Then
    Call Library.setComment(Library.getRegistry("Main", "CommentBgColor") _
                          , Library.getRegistry("Main", "CommentFont") _
                          , Library.getRegistry("Main", "CommentFontColor") _
                          , Library.getRegistry("Main", "CommentFontSize") _
                          )
    
  End If
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ړ���T�C�Y�ύX������()
  Dim slctObect
  Const funcName As String = "Ctl_Format.�ړ���T�C�Y�ύX������"
 
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        slctObect.Placement = xlMoveAndSize
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlMoveAndSize
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ړ�����()
  Dim slctObect
  Const funcName As String = "Ctl_Format.�ړ�����"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        slctObect.Placement = xlMove
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlMove
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ړ���T�C�Y�ύX�����Ȃ�()
  Dim slctObect
  Const funcName As String = "Ctl_Format.�ړ���T�C�Y�ύX�����Ȃ�"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        slctObect.Placement = xlFreeFloating
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlFreeFloating
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �]���[��()
  Dim slctObect
  Const funcName As String = "Ctl_Format.�]���[��"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", CStr(runFlg), "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle"
        Call Library.showDebugForm("TypeName", "Set�F" & TypeName(slctObect), "debug")
        
        slctObect.ShapeRange.TextFrame2.MarginTop = 0
        slctObect.ShapeRange.TextFrame2.MarginBottom = 0
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
        
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


