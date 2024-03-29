Attribute VB_Name = "Ctl_Stamp"
Option Explicit

'==================================================================================================
Function Ïó()
  
  Const funcName As String = "Ctl_Stamp.Ïó"
  
  'Jn--------------------------------------
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
  
  Set targetSheet = ActiveSheet
  Set targetRange = ActiveCell
  
  LadexSh_Stamp.Activate
  LadexSh_Stamp.Shapes.Range(Array("Ïó")).Select
  Selection.copy
  
  targetSheet.Activate
  ActiveSheet.PasteSpecial Format:="} (PNG)", Link:=False, DisplayAsIcon:=False
  Selection.Placement = xlMoveAndSize
  Selection.ShapeRange.LockAspectRatio = msoFalse
  
  Selection.ShapeRange.Width = 30
  Selection.ShapeRange.Height = 30
  Selection.ShapeRange.Name = "Ïó_" & Format(Now(), "yyyymmdd_hhnnss")
  
  targetRange.Select
  
  'I¹--------------------------------------
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

  'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function mFó(Optional StampName As String, Optional StampVal As String, Optional StampFont As String, Optional imgName As String)
  
  Const funcName As String = "Ctl_Stamp.mFó"
  
  'Jn--------------------------------------
  If runFlg = False Then
    Call init.setting
    Call Library.showDebugForm(funcName, , "start")
    Call Library.startScript
    Call Ctl_ProgressBar.showStart
    PrgP_Max = 2
  Else
'    On Error GoTo catchError
    Call Library.showDebugForm(funcName, , "start1")
  End If
  Call Library.showDebugForm("runFlg", runFlg, "debug")
  PrgP_Cnt = PrgP_Cnt + 1
  '----------------------------------------------
  
  Set targetSheet = ActiveSheet
  Set targetRange = ActiveCell
  
  If StampName = "" Then
    StampName = Library.getRegistry("Main", "StampName")
    StampFont = Library.getRegistry("Main", "StampFont")
    StampVal = Library.getRegistry("Main", "StampVal")
  End If
  
  LadexSh_Stamp.Activate
  LadexSh_Stamp.Shapes.Range(Array("TB_Ï")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = StampVal
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameComplexScript = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameFarEast = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Name = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Size = Ctl_shap.TextToFitShape(Selection.ShapeRange(1), True)
  
  LadexSh_Stamp.Shapes.Range(Array("TB_¼O1")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = StampName
  Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 80
  
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameComplexScript = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameFarEast = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Name = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Size = Ctl_shap.TextToFitShape(Selection.ShapeRange(1), False)

  Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
  
  
  
  LadexSh_Stamp.Shapes.Range(Array("TB_út")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(Now(), "yyyy/m/d")
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameComplexScript = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameFarEast = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Name = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Size = Ctl_shap.TextToFitShape(Selection.ShapeRange(1), True)
  
  
  LadexSh_Stamp.Shapes.Range(Array("mFó")).Select
  Selection.copy
  
  targetSheet.Activate
  ActiveSheet.PasteSpecial Format:="} (PNG)", Link:=False, DisplayAsIcon:=False
  Selection.Placement = xlMoveAndSize
  Selection.ShapeRange.LockAspectRatio = msoFalse
  
  Selection.ShapeRange.Width = 75
  Selection.ShapeRange.Height = 75
  
  If imgName = "" Then
    Selection.ShapeRange.Name = "mFó_" & Format(Now(), "yyyymmdd_hhnnss")
  Else
    Selection.ShapeRange.Name = imgName
  End If
    
  targetRange.Select
  
  'I¹--------------------------------------
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

  'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function


'==================================================================================================
Function ¼O()
  Dim StampName As String, StampFont As String
  Const funcName As String = "Ctl_Stamp.¼O"
  
  'Jn--------------------------------------
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
  
  Set targetSheet = ActiveSheet
  Set targetRange = ActiveCell
  
  'StampFont = Library.getRegistry("Main", "StampFont")
  StampFont = "HGSsÌ"
  StampName = Library.getRegistry("Main", "StampName")
  
  
  LadexSh_Stamp.Activate
  LadexSh_Stamp.Shapes.Range(Array("TB_¼O2")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Split(StampName, " ")(0)
  'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = StampName
  
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameComplexScript = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.NameFarEast = StampFont
  Selection.ShapeRange.TextFrame2.TextRange.Font.Name = StampFont
  
  Selection.ShapeRange.TextFrame2.TextRange.Font.Size = Ctl_shap.TextToFitShape(Selection.ShapeRange(1))

  
  LadexSh_Stamp.Shapes.Range(Array("Fó")).Select
  Selection.copy
  
  targetSheet.Activate
  ActiveSheet.PasteSpecial Format:="} (PNG)", Link:=False, DisplayAsIcon:=False
  Selection.Placement = xlMoveAndSize
  Selection.ShapeRange.LockAspectRatio = msoFalse
  
  Selection.ShapeRange.Width = 55
  Selection.ShapeRange.Height = 55
  Selection.ShapeRange.Name = "Fó_" & Format(Now(), "yyyymmdd_hhnnss")
    
  targetRange.Select
  
  'I¹--------------------------------------
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

  'G[­¶------------------------------------
catchError:
  Call Library.showNotice(400, "<" & funcName & "[" & Err.Number & "]" & Err.Description & ">", True)
End Function
