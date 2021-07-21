Attribute VB_Name = "Main"
'ワークブック用変数------------------------------
'ワークシート用変数------------------------------
'グローバル変数----------------------------------

'==================================================================================================
Function xxxxxxxxxx()
End Function





'**************************************************************************************************
' * 標準画面
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 標準画面()
  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  
  'On Error Resume Next
  
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  SelectAddress = Selection.Address
  
  setZoomLevel = Library.getRegistry("Main", "zoomLevel")
  resetBgColor = Library.getRegistry("Main", "bgColor")
  setGgridLine = Library.getRegistry("Main", "gridLine")
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    
    Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 0, 4, sheetName)
    
    If Worksheets(sheetName).Visible = True Then
      Worksheets(sheetName).Select
      
      '標準画面に設定
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 1, 4, sheetName)
      ActiveWindow.View = xlNormalView
      
      '表示倍率の指定
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 2, 4, sheetName)
      ActiveWindow.Zoom = setZoomLevel
      
      'ガイドラインの表示/非表示
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 3, 4, sheetName)
      ActiveWindow.DisplayGridlines = setGgridLine
  
      '背景白をなしにする
      Call Ctl_ProgressBar.showBar("標準画面設定", sheetCount, sheetMaxCount, 4, 4, sheetName)
      If resetBgColor = True Then
        With Application.FindFormat.Interior
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        With Application.ReplaceFormat.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
      End If
      Application.GoTo Reference:=Range("A1"), Scroll:=True
    End If
    
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  Range(SelectAddress).Select
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  
End Function

'==================================================================================================
Function A1セル選択()

  Dim objSheet As Object
  Dim sheetName As String, SetActiveSheet As String
  Dim sheetCount As Long, sheetMaxCount As Long
  
  'On Error Resume Next
  
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  
  SetActiveSheet = ActiveWorkbook.ActiveSheet.Name
  
  sheetCount = 0
  sheetMaxCount = ActiveWorkbook.Sheets.count
  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name
    
    If Worksheets(sheetName).Visible = True Then
      Worksheets(sheetName).Select
      Application.GoTo Reference:=Range("A1"), Scroll:=True
    End If
    
    sheetCount = sheetCount + 1
  Next
  
  Worksheets(SetActiveSheet).Select
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  
End Function
'==================================================================================================




'**************************************************************************************************
' * 名前定義削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義削除()
  Dim wb As Workbook, tmp As String
  
  Call Library.startScript
  
  For Each wb In Workbooks
    Workbooks(wb.Name).Activate
    Call Library.delVisibleNames
  Next wb
  
  Call Library.endScript

End Function


'**************************************************************************************************
' * 画像設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 画像設定()

  With ActiveWorkbook.ActiveSheet
    Dim AllShapes As Shapes
    Dim CurShape As Shape
    Set AllShapes = .Shapes
    
    For Each CurShape In AllShapes
      CurShape.Placement = xlMove
    Next
  End With
  
End Function



'**************************************************************************************************
' * ハイライト
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ハイライト()
'  Dim highLightFlg As String
'  Dim highLightArea As String
'
'  Call Library.startScript
'  highLightFlg = Library.getRegistry(ActiveWorkbook.Name, "HighLightFlg")
'
'  If highLightFlg = "" Then
'    Call Library.setLineColor(Selection.Address, True, Library.getRegistry("HighLightColor"))
'
'    Call Library.setRegistry(ActiveWorkbook.Name, True, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightSheet", ActiveSheet.Name, "HighLightFlg")
'    Call Library.setRegistry(ActiveWorkbook.Name & "_HighLightArea", Selection.Address, "HighLightFlg")
'
'  Else
'    highLightArea = Library.getRegistry(ActiveWorkbook.Name & "_HighLightArea")
'
'    If highLightArea = "" Then
'      highLightArea = Selection.Address
'    End If
'    Call Library.unsetLineColor(highLightArea)
'
'    Call Library.delRegistry(ActiveWorkbook.Name, "HighLightFlg")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightSheet")
'    Call Library.delRegistry(ActiveWorkbook.Name & "_HighLightArea")
'  End If
'
'  Call Library.endScript(True)

End Function

'**************************************************************************************************
' * コメント整形
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function コメント整形()
  
  On Error GoTo catchError
  
  If TypeName(ActiveCell) = "Range" Then
    Call Library.setComment(Library.getRegistry("Main", "CommentBgColor"))
  
  End If
  
  
  Exit Function

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function



'**************************************************************************************************
' * すべて表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function すべて表示()
  Dim rowOutlineLevel As Long, colOutlineLevel As Long
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting

  If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
  End If
  If ActiveWindow.DisplayOutline = True Then
    For rowOutlineLevel = 1 To 15
      DoEvents
      ActiveSheet.Outline.ShowLevels rowLevels:=rowOutlineLevel
      If Err.Number <> 0 Then
        Err.Clear
        Exit For
      End If
    Next
    
    For colOutlineLevel = 1 To 15
      DoEvents
      ActiveSheet.Outline.ShowLevels columnLevels:=colOutlineLevel
      If Err.Number <> 0 Then
        Err.Clear
        Exit For
      End If
    Next
  End If
  ActiveSheet.Cells.EntireColumn.Hidden = False
  ActiveSheet.Cells.EntireRow.Hidden = False
  
  Call Library.endScript(, True)
End Function


'**************************************************************************************************
' * 設定Import / Export
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function 設定_抽出()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting
  
  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  BK_sheetStyle.Copy
  ActiveWorkbook.SaveAs fileName:=TempName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  
  Call Library.endScript
  
  MsgBox ("修正完了後、保存し閉じてください")
End Function

'==================================================================================================
Function 設定_取込()
  
  Dim FSO As Object, TempName As String
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  Call Library.startScript
  Call init.setting

  TempName = FSO.GetSpecialFolder(2) & "\BK_Style.xlsx"
  
  Set targetBook = Workbooks.Open(TempName)
  targetBook.Sheets("Style").Columns("A:J").Copy ThisWorkbook.Worksheets("Style").Range("A1")
  targetBook.Close
  
  Call FSO.DeleteFile(TempName, True)
  
  Call Main.スタイル削除
  Call Library.endScript
End Function











