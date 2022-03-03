Attribute VB_Name = "Ctl_format"
Option Explicit

'**************************************************************************************************
' * コメント整形
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function コメント整形()
  Const funcName As String = "Ctl_Format.コメント整形"
  
  '処理開始--------------------------------------
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
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 移動やサイズ変更をする()
  Dim slctObect
  Const funcName As String = "Ctl_Format.移動やサイズ変更をする"
 
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        slctObect.Placement = xlMoveAndSize
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlMoveAndSize
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 移動する()
  Dim slctObect
  Const funcName As String = "Ctl_Format.移動する"
  
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        slctObect.Placement = xlMove
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlMove
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 移動やサイズ変更をしない()
  Dim slctObect
  Const funcName As String = "Ctl_Format.移動やサイズ変更をしない"
  
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        slctObect.Placement = xlFreeFloating
      
      Case "ChartObject"
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        
        Call Library.showDebugForm("ShapeRange.name", slctObect.ShapeRange.Name, "debug")
        ActiveSheet.ChartObjects(slctObect.ShapeRange.Name).Activate
        Selection.Placement = xlFreeFloating
      
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 余白ゼロ()
  Dim slctObect
  Const funcName As String = "Ctl_Format.余白ゼロ"
  
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "Set：" & TypeName(slctObect), "debug")
        
        slctObect.ShapeRange.TextFrame2.MarginTop = 0
        slctObect.ShapeRange.TextFrame2.MarginBottom = 0
      Case Else
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
        
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm("", , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm("", , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


