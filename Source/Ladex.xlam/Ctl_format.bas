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

'==================================================================================================
Function �����T�C�Y���҂�����ɂ���()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Format.�����T�C�Y���҂�����ɂ���"

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
  
  Call Ctl_shap.TextToFitShape(Selection.ShapeRange(1), True)


  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function �Z�����̒����ɔz�u()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range, targetRange As Range
  Dim ShapeImg As Shape
  
  Const funcName As String = "Ctl_Format.�Z�����̒����ɔz�u"

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
   
  For Each slctCells In Selection
    For Each ShapeImg In ActiveSheet.Shapes
      Set targetRange = Range(ShapeImg.TopLeftCell, ShapeImg.BottomRightCell)
      If Not (Intersect(targetRange, slctCells) Is Nothing) Then
        Call Library.showDebugForm("ShapeImg.Name  ", ShapeImg.Name, "debug")
        Call Library.showDebugForm("ShapeImg.Width  ", ShapeImg.Width, "debug")
        Call Library.showDebugForm("ShapeImg.Height ", ShapeImg.Height, "debug")
        Call Library.showDebugForm("slctCells.Width ", slctCells.Width, "debug")
        Call Library.showDebugForm("slctCells.Height", slctCells.Height, "debug")
        
        With ShapeImg
          .Top = slctCells.Top + (slctCells.Height - ShapeImg.Height) / 2
          .Left = slctCells.Left + (slctCells.Width - ShapeImg.Width) / 2
        End With
        
      End If
    Next
  Next

  '�����I��--------------------------------------
  If runFlg = False Then
    Application.GoTo Reference:=Range("A1"), Scroll:=True
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
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function
