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
  Call Library.showDebugForm("runFlg", runFlg, "debug")
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
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
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
  Dim slctObect As Variant
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
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

'==================================================================================================
Function 上下余白ゼロ()
  Dim slctObect
  Const funcName As String = "Ctl_Format.上下余白ゼロ"
  
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
        
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Function 左右余白ゼロ()
  Dim slctObect
  Const funcName As String = "Ctl_Format.左右余白ゼロ"
  
  '処理開始--------------------------------------
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
        Call Library.showDebugForm("TypeName", "NotSet：" & TypeName(slctObect), "debug")
        Call Library.showError("対象となるシェイプではありません")
        
    End Select
  Next
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
    Call Library.showDebugForm(funcName, , "end")
    Call init.unsetting
  Else
    Call Library.showDebugForm(funcName, , "end")
  End If
  '----------------------------------------------
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function

