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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
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
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection.ShapeRange
    Call Library.showDebugForm("TypeName  ", TypeName(slctObect), "debug")
    Call Library.showDebugForm("ObjName ", slctObect.Name, "debug")
    
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        slctObect.Placement = xlMoveAndSize
      
      Case "ChartObject"
        ActiveSheet.ChartObjects(slctObect.Name).Activate
        Selection.Placement = xlMoveAndSize
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection.ShapeRange
    Call Library.showDebugForm("TypeName", TypeName(slctObect), "debug")
    Call Library.showDebugForm("ObjName ", slctObect.Name, "debug")
    
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        slctObect.Placement = xlMove
      
      Case "ChartObject"
        ActiveSheet.ChartObjects(slctObect.Name).Activate
        Selection.Placement = xlMove
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
  Dim slctObect As Variant
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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection.ShapeRange
    Call Library.showDebugForm("TypeName  ", TypeName(slctObect), "debug")
    Call Library.showDebugForm("ObjName ", slctObect.Name, "debug")

    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Picture", "Shape"
        slctObect.Placement = xlFreeFloating
      
      Case "ChartObject"
        ActiveSheet.ChartObjects(slctObect.Name).Activate
        Selection.Placement = xlFreeFloating
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �㉺�]���[��()
  Dim slctObect
  Const funcName As String = "Ctl_Format.�㉺�]���[��"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection.ShapeRange
    Call Library.showDebugForm("TypeName  ", TypeName(slctObect), "debug")
    Call Library.showDebugForm("ObjName ", slctObect.Name, "debug")
    
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Shape"
        
        slctObect.TextFrame2.MarginTop = 0
        slctObect.TextFrame2.MarginBottom = 0
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
        
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function ���E�]���[��()
  Dim slctObect
  Const funcName As String = "Ctl_Format.���E�]���[��"
  
  '�����J�n--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Else
    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  '----------------------------------------------
  
  For Each slctObect In Selection.ShapeRange
    Call Library.showDebugForm("TypeName  ", TypeName(slctObect), "debug")
    Call Library.showDebugForm("ObjName ", slctObect.Name, "debug")
    
    Select Case TypeName(slctObect)
      Case "TextBox", "Rectangle", "Shape"
        
        slctObect.TextFrame2.MarginLeft = 0
        slctObect.TextFrame2.MarginRight = 0
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet�F" & TypeName(slctObect), "debug")
        Call Library.showError("�ΏۂƂȂ�V�F�C�v�ł͂���܂���")
        
    End Select
  Next
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

