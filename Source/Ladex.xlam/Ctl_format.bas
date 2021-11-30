Attribute VB_Name = "Ctl_format"
Option Explicit

'**************************************************************************************************
' * �R�����g���`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �R�����g���`()
  
  On Error GoTo catchError
  Call init.setting
  
  If TypeName(ActiveCell) = "Range" Then
    Call Library.startScript
    Call Library.setComment(Library.getRegistry("Main", "CommentBgColor") _
                          , Library.getRegistry("Main", "CommentFont") _
                          , Library.getRegistry("Main", "CommentFontColor") _
                          , Library.getRegistry("Main", "CommentFontSize") _
                          )
    
    Call Library.endScript
  End If
  
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function �ړ���T�C�Y�ύX������()
  
  On Error GoTo catchError
  Call init.setting
  
  Select Case TypeName(Selection)
    Case "TextBox", "Rectangle", "Picture"
      Selection.Placement = xlMoveAndSize
  End Select
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ړ�����()
  
  On Error GoTo catchError
  Call init.setting
  
  Select Case TypeName(Selection)
    Case "TextBox", "Rectangle", "Picture"
      Selection.Placement = xlMove
  End Select
  
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �ړ���T�C�Y�ύX�����Ȃ�()
  Const funcName As String = "Ctl_Format.�ړ���T�C�Y�ύX�����Ȃ�"
  
  On Error GoTo catchError
  Call init.setting
  
  Select Case TypeName(Selection)
    Case "TextBox", "Rectangle", "Picture"
      Selection.Placement = xlFreeFloating
    Case "ChartArea"
    
    
  End Select
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function �]���[��()
  
  On Error GoTo catchError
  Call init.setting
  
  Select Case TypeName(Selection)
    Case "TextBox"
      Selection.ShapeRange.TextFrame2.MarginTop = 0
      Selection.ShapeRange.TextFrame2.MarginBottom = 0
    '  Selection.ShapeRange.TextFrame2.MarginLeft = 0
    '  Selection.ShapeRange.TextFrame2.MarginRight = 0
  End Select
  
  
  Exit Function

'�G���[������------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


