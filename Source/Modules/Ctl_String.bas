Attribute VB_Name = "Ctl_String"
Option Explicit

'**************************************************************************************************
' * 文字列操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Trim01()
  
  Call init.setting
  ActiveCell = Trim(ActiveCell.Text)
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function

'==================================================================================================
Function 中黒点付与()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.中黒点付与"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  For Each slctCells In Selection
    slctCells.Value = "・" & slctCells.Value
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function 連番設定()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.連番設定"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
    slctCells.Value = line & "．"
    line = line + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function



'==================================================================================================
Function 連番追加()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.連番追加"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  line = 1
  For Each slctCells In Selection
    slctCells.Value = line & "." & slctCells.Value
    line = line + 1
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function



'==================================================================================================
Function 英数字全半角変換()
  Dim line As Long, endLine As Long
  Dim slctCells As Range
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_String.英数字全半角変換"

  Call Library.startScript
  Call init.setting
  '----------------------------------------------
  For Each slctCells In Selection
    slctCells.Value = Library.convHan2Zen(slctCells.Value)
    DoEvents
  Next

  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

