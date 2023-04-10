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

'==================================================================================================
Function 文字サイズをぴったりにする()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Const funcName As String = "Ctl_Format.文字サイズをぴったりにする"

  '処理開始--------------------------------------
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


  '処理終了--------------------------------------
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

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function セル内の中央に配置()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim slctCells As Range, targetRange As Range
  Dim ShapeImg As Shape
  
  Const funcName As String = "Ctl_Format.セル内の中央に配置"

  '処理開始--------------------------------------
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

  '処理終了--------------------------------------
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

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function
