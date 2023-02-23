Attribute VB_Name = "Ctl_Option"
'**************************************************************************************************
' * オプションフォーム制御
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function オプション画面表示()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim images As Variant, tmpObjChart As Variant
  Dim CompanyHolidayList As String, dataExtractList As String
  
'  On Error GoTo catchError
  
  
  With Frm_Option
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    
    'マルチページの表示
    .MultiPage1.value = 0
    
    '期間、基準日の初期値
'    .startDay.Text = setVal("startDay")
'    .endDay.Text = setVal("GUNT_END_DAY")
    
    .startDay.Text = Format(Date, "yyyy/mm/dd")
    .endDay.Text = DateSerial(Year(Date), Month(Date) + 3, 0)
    
    .baseDay.Text = setVal("baseDay")
    
    .setLightning.value = setVal("setLightning")
    .setDispProgress100.value = setVal("setDispProgress100")
        
    'スタイル関連
    .lineColor.BackColor = setVal("lineColor")
    .SaturdayColor.BackColor = setVal("SaturdayColor")
    .SundayColor.BackColor = setVal("SundayColor")
    .CompanyHolidayColor.BackColor = setVal("CompanyHolidayColor")
    .lineColor_Plan.BackColor = setVal("lineColor_Plan")
    .lineColor_Achievement.BackColor = setVal("lineColor_Achievement")
    .lineColor_Lightning.BackColor = setVal("lineColor_Lightning")
    .lineColor_TaskLevel1.BackColor = setVal("lineColor_TaskLevel1")
    .lineColor_TaskLevel2.BackColor = setVal("lineColor_TaskLevel2")
    .lineColor_TaskLevel3.BackColor = setVal("lineColor_TaskLevel3")
    
    
    'ショートカットキー関連
    .optionKey.value = setVal("optionKey")
    .centerKey.value = setVal("centerKey")
    .filterKey.value = setVal("filterKey")
    .clearFilterKey.value = setVal("clearFilterKey")
    .taskCheckKey.value = setVal("taskCheckKey")
    .makeGanttKey.value = setVal("makeGanttKey")
    .clearGanttKey.value = setVal("clearGanttKey")
    .dispAllKey.value = setVal("dispAllKey")
    .taskControlKey.value = setVal("taskControlKey")
    .ScaleKey.value = setVal("ScaleKey")
    
    '担当者
    .Assign01.Text = sh_Option.Range(setVal("cell_AssignorList") & 4)
    .Assign02.Text = sh_Option.Range(setVal("cell_AssignorList") & 5)
    .Assign03.Text = sh_Option.Range(setVal("cell_AssignorList") & 6)
    .Assign04.Text = sh_Option.Range(setVal("cell_AssignorList") & 7)
    .Assign05.Text = sh_Option.Range(setVal("cell_AssignorList") & 8)
    .Assign06.Text = sh_Option.Range(setVal("cell_AssignorList") & 9)
    .Assign07.Text = sh_Option.Range(setVal("cell_AssignorList") & 10)
    .Assign08.Text = sh_Option.Range(setVal("cell_AssignorList") & 11)
    .Assign09.Text = sh_Option.Range(setVal("cell_AssignorList") & 12)
    .Assign10.Text = sh_Option.Range(setVal("cell_AssignorList") & 13)
    .Assign11.Text = sh_Option.Range(setVal("cell_AssignorList") & 14)
    .Assign12.Text = sh_Option.Range(setVal("cell_AssignorList") & 15)
    .Assign13.Text = sh_Option.Range(setVal("cell_AssignorList") & 16)
    .Assign14.Text = sh_Option.Range(setVal("cell_AssignorList") & 17)
    .Assign15.Text = sh_Option.Range(setVal("cell_AssignorList") & 18)
    .Assign16.Text = sh_Option.Range(setVal("cell_AssignorList") & 19)
    .Assign17.Text = sh_Option.Range(setVal("cell_AssignorList") & 20)
    .Assign18.Text = sh_Option.Range(setVal("cell_AssignorList") & 21)
    .Assign19.Text = sh_Option.Range(setVal("cell_AssignorList") & 22)
    .Assign20.Text = sh_Option.Range(setVal("cell_AssignorList") & 23)
    .Assign21.Text = sh_Option.Range(setVal("cell_AssignorList") & 24)
    .Assign22.Text = sh_Option.Range(setVal("cell_AssignorList") & 25)
    .Assign23.Text = sh_Option.Range(setVal("cell_AssignorList") & 26)
    .Assign24.Text = sh_Option.Range(setVal("cell_AssignorList") & 27)
    .Assign25.Text = sh_Option.Range(setVal("cell_AssignorList") & 28)
    .Assign26.Text = sh_Option.Range(setVal("cell_AssignorList") & 29)
    .Assign27.Text = sh_Option.Range(setVal("cell_AssignorList") & 30)
    .Assign28.Text = sh_Option.Range(setVal("cell_AssignorList") & 31)
    .Assign29.Text = sh_Option.Range(setVal("cell_AssignorList") & 32)
    .Assign30.Text = sh_Option.Range(setVal("cell_AssignorList") & 33)
    .Assign31.Text = sh_Option.Range(setVal("cell_AssignorList") & 34)
    .Assign32.Text = sh_Option.Range(setVal("cell_AssignorList") & 35)
    .Assign33.Text = sh_Option.Range(setVal("cell_AssignorList") & 36)
    .Assign34.Text = sh_Option.Range(setVal("cell_AssignorList") & 37)
    .Assign35.Text = sh_Option.Range(setVal("cell_AssignorList") & 38)
    
    '担当者色
    .AssignColor01.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 4).Interior.Color
    .AssignColor02.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 5).Interior.Color
    .AssignColor03.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 6).Interior.Color
    .AssignColor04.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 7).Interior.Color
    .AssignColor05.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 8).Interior.Color
    .AssignColor06.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 9).Interior.Color
    .AssignColor07.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 10).Interior.Color
    .AssignColor08.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 11).Interior.Color
    .AssignColor09.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 12).Interior.Color
    .AssignColor10.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 13).Interior.Color
    .AssignColor11.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 14).Interior.Color
    .AssignColor12.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 15).Interior.Color
    .AssignColor13.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 16).Interior.Color
    .AssignColor14.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 17).Interior.Color
    .AssignColor15.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 18).Interior.Color
    .AssignColor16.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 19).Interior.Color
    .AssignColor17.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 20).Interior.Color
    .AssignColor18.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 21).Interior.Color
    .AssignColor19.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 22).Interior.Color
    .AssignColor20.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 23).Interior.Color
    .AssignColor21.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 24).Interior.Color
    .AssignColor22.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 25).Interior.Color
    .AssignColor23.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 26).Interior.Color
    .AssignColor24.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 27).Interior.Color
    .AssignColor25.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 28).Interior.Color
    .AssignColor26.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 29).Interior.Color
    .AssignColor27.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 30).Interior.Color
    .AssignColor28.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 31).Interior.Color
    .AssignColor29.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 32).Interior.Color
    .AssignColor30.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 33).Interior.Color
    .AssignColor31.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 34).Interior.Color
    .AssignColor32.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 35).Interior.Color
    .AssignColor33.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 36).Interior.Color
    .AssignColor34.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 37).Interior.Color
    .AssignColor35.BackColor = sh_Option.Range(setVal("cell_AssignorList") & 38).Interior.Color
    
    '担当者単価
    .unitCost01.Text = sh_Option.Range(setVal("cell_unitCostorList") & 4)
    .unitCost02.Text = sh_Option.Range(setVal("cell_unitCostorList") & 5)
    .unitCost03.Text = sh_Option.Range(setVal("cell_unitCostorList") & 6)
    .unitCost04.Text = sh_Option.Range(setVal("cell_unitCostorList") & 7)
    .unitCost05.Text = sh_Option.Range(setVal("cell_unitCostorList") & 8)
    .unitCost06.Text = sh_Option.Range(setVal("cell_unitCostorList") & 9)
    .unitCost07.Text = sh_Option.Range(setVal("cell_unitCostorList") & 10)
    .unitCost08.Text = sh_Option.Range(setVal("cell_unitCostorList") & 11)
    .unitCost09.Text = sh_Option.Range(setVal("cell_unitCostorList") & 12)
    .unitCost10.Text = sh_Option.Range(setVal("cell_unitCostorList") & 13)
    .unitCost11.Text = sh_Option.Range(setVal("cell_unitCostorList") & 14)
    .unitCost12.Text = sh_Option.Range(setVal("cell_unitCostorList") & 15)
    .unitCost13.Text = sh_Option.Range(setVal("cell_unitCostorList") & 16)
    .unitCost14.Text = sh_Option.Range(setVal("cell_unitCostorList") & 17)
    .unitCost15.Text = sh_Option.Range(setVal("cell_unitCostorList") & 18)
    .unitCost16.Text = sh_Option.Range(setVal("cell_unitCostorList") & 19)
    .unitCost17.Text = sh_Option.Range(setVal("cell_unitCostorList") & 20)
    .unitCost18.Text = sh_Option.Range(setVal("cell_unitCostorList") & 21)
    .unitCost19.Text = sh_Option.Range(setVal("cell_unitCostorList") & 22)
    .unitCost20.Text = sh_Option.Range(setVal("cell_unitCostorList") & 23)
    .unitCost21.Text = sh_Option.Range(setVal("cell_unitCostorList") & 24)
    .unitCost22.Text = sh_Option.Range(setVal("cell_unitCostorList") & 25)
    .unitCost23.Text = sh_Option.Range(setVal("cell_unitCostorList") & 26)
    .unitCost24.Text = sh_Option.Range(setVal("cell_unitCostorList") & 27)
    .unitCost25.Text = sh_Option.Range(setVal("cell_unitCostorList") & 28)
    .unitCost26.Text = sh_Option.Range(setVal("cell_unitCostorList") & 29)
    .unitCost27.Text = sh_Option.Range(setVal("cell_unitCostorList") & 30)
    .unitCost28.Text = sh_Option.Range(setVal("cell_unitCostorList") & 31)
    .unitCost29.Text = sh_Option.Range(setVal("cell_unitCostorList") & 32)
    .unitCost30.Text = sh_Option.Range(setVal("cell_unitCostorList") & 33)
    .unitCost31.Text = sh_Option.Range(setVal("cell_unitCostorList") & 34)
    .unitCost32.Text = sh_Option.Range(setVal("cell_unitCostorList") & 35)
    .unitCost33.Text = sh_Option.Range(setVal("cell_unitCostorList") & 36)
    .unitCost34.Text = sh_Option.Range(setVal("cell_unitCostorList") & 37)
    .unitCost35.Text = sh_Option.Range(setVal("cell_unitCostorList") & 38)

    
    '会社指定休日
    For line = 3 To sh_Option.Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).Row
      If sh_Option.Range(setVal("cell_CompanyHoliday") & line) <> "" Then
        If CompanyHolidayList = "" Then
          CompanyHolidayList = sh_Option.Range(setVal("cell_CompanyHoliday") & line)
        Else
          CompanyHolidayList = CompanyHolidayList & vbNewLine & sh_Option.Range(setVal("cell_CompanyHoliday") & line)
        End If
      End If
    Next
    .CompanyHoliday.Text = CompanyHolidayList
    
    '抽出タスク
    For line = 3 To sh_Option.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).Row
      If sh_Option.Range(setVal("cell_DataExtract") & line) <> "" Then
        If dataExtractList = "" Then
          dataExtractList = sh_Option.Range(setVal("cell_DataExtract") & line)
        Else
          dataExtractList = dataExtractList & vbNewLine & sh_Option.Range(setVal("cell_DataExtract") & line)
        End If
      End If
    Next
    .dataExtract.Text = dataExtractList
    
    
    '表示設定
    .view_Plan.value = setVal("view_Plan")
    .view_Assign.value = setVal("view_Assign")
    .view_Progress.value = setVal("view_Progress")
    .view_Achievement.value = setVal("view_Achievement")
    .view_Task.value = setVal("view_Task")
    .view_TaskInfo.value = setVal("view_TaskInfo")
    .view_TaskAllocation.value = setVal("view_TaskAllocation")
    .view_LineInfo.value = setVal("view_LineInfo")
    
    .view_WorkLoad.value = setVal("view_WorkLoad")
    .view_LateOrEarly.value = setVal("view_LateOrEarly")
    .view_Note.value = setVal("view_Note")
    
    .viewGant_TaskName.value = setVal("viewGant_TaskName")
    .viewGant_Assignor.value = setVal("viewGant_Assignor")
  
  End With
  
  Frm_Option.Show

  Exit Function
'エラー発生時------------------------------------

catchError:

End Function


'==================================================================================================
Function オプション設定値格納()

  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim CompanyHoliday As Variant, dataExtract As Variant

  On Error Resume Next
  
  Call Ctl_ProgressBar.showStart
  sh_Option.Select
'  For line = 3 To sh_Option.Range("B5")
'    'Call Ctl_ProgressBar.showCount("オプション設定値格納", line, sh_Option.Range("B5"), sh_Option.Range("A" & line) & ":" & getVal(sh_Option.Range("A" & line)))
'    sh_Option.Range(sh_Option.Range("A" & line)).Select
'
'    If IsEmpty(getVal(sh_Option.Range("A" & line))) = False Then
'      Select Case sh_Option.Range("A" & line)
'        Case "baseDay"
'          If getVal(sh_Option.Range("A" & line)) = Format(Now, "yyyy/mm/dd") Then
'            sh_Option.Range(sh_Option.Range("A" & line)).FormulaR1C1 = "=TODAY()"
'          Else
'            sh_Option.Range(sh_Option.Range("A" & line)) = getVal(sh_Option.Range("A" & line))
'          End If
'
'        Case ""
'        Case Else
'          sh_Option.Range(sh_Option.Range("A" & line)) = getVal(sh_Option.Range("A" & line))
'      End Select
'    End If
'  Next
  
'  'ショートカットキーの設定
'  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
'  For line = 3 To endLine
'    Call Ctl_ProgressBar.showCount("オプション設定値格納", line, sh_Option.Range("B5"), "ショートカットキー設定")
'
'    Range(Range(setVal("cell_ShortcutFuncName") & line)).Select
'    Range(Range(setVal("cell_ShortcutFuncName") & line)) = getVal(Range(setVal("cell_ShortcutFuncName") & line))
'  Next
'
'  '会社指定休日の設定
'  line = 3
'  sh_Option.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).row).ClearContents
'  For Each CompanyHoliday In Split(getVal("CompanyHoliday"), vbNewLine)
'    DoEvents
'    sh_Option.Range(setVal("cell_CompanyHoliday") & line) = CompanyHoliday
'    line = line + 1
'  Next
'
'  '抽出タスクの設定
'  line = 3
'  sh_Option.Range(setVal("cell_DataExtract") & "3:" & setVal("cell_DataExtract") & Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row).ClearContents
'  For Each dataExtract In Split(getVal("dataExtract"), vbNewLine)
'    Call Ctl_ProgressBar.showCount("オプション設定値格納", line, 100, "抽出タスクの設定")
'
'    sh_Option.Range(setVal("cell_DataExtract") & line) = dataExtract
'    line = line + 1
'  Next


  '担当者
  sh_Option.Range(setVal("cell_AssignorList") & "4:" & setVal("cell_AssignorList") & Cells(Rows.count, Library.getColumnNo(setVal("cell_AssignorList"))).End(xlUp).Row).Clear
  For line = 4 To 38
    'Call Ctl_ProgressBar.showCount("オプション設定値格納", line, 38, "担当者:" & getVal("Assign" & Format(line - 3, "00")))
    
    sh_Option.Range(setVal("cell_AssignorList") & line) = getVal("Assign" & Format(line - 3, "00"))
    sh_Option.Range(setVal("cell_AssignorList") & line).Interior.Color = getVal("AssignColor" & Format(line - 3, "00"))
    
    sh_Option.Range(setVal("cell_unitCostorList") & line) = getVal("unitCost" & Format(line - 3, "00"))
    
  Next
  sh_Option.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & 38).Select

'  Call 罫線.囲み罫線
'  Call Menu.M_ショートカット設定
  
  ActiveWorkbook.Worksheets("WBS").Select
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  

  Call Ctl_ProgressBar.showEnd
  Set getVal = Nothing
  
End Function
