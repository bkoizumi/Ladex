Attribute VB_Name = "Ctl_Option"
'**************************************************************************************************
' * �I�v�V�����t�H�[������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �I�v�V������ʕ\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim images As Variant, tmpObjChart As Variant
  Dim CompanyHolidayList As String, dataExtractList As String
  
'  On Error GoTo catchError
  
  Call WBS_Option.�ݒ�V�[�g�R�s�[("forAddin")
  With Frm_Option
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    
    '�}���`�y�[�W�̕\��
    .MultiPage1.Value = 0
    
    '���ԁA����̏����l
'    .startDay.Text = setVal("startDay")
'    .endDay.Text = setVal("GUNT_END_DAY")
    
    .startDay.Text = Format(Date, "yyyy/mm/dd")
    .endDay.Text = DateSerial(Year(Date), Month(Date) + 3, 0)
    
    .baseDay.Text = setVal("baseDay")
    
    .setLightning.Value = setVal("setLightning")
    .setDispProgress100.Value = setVal("setDispProgress100")
        
    '�X�^�C���֘A
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
    
    
    '�V���[�g�J�b�g�L�[�֘A
    .optionKey.Value = setVal("optionKey")
    .centerKey.Value = setVal("centerKey")
    .filterKey.Value = setVal("filterKey")
    .clearFilterKey.Value = setVal("clearFilterKey")
    .taskCheckKey.Value = setVal("taskCheckKey")
    .makeGanttKey.Value = setVal("makeGanttKey")
    .clearGanttKey.Value = setVal("clearGanttKey")
    .dispAllKey.Value = setVal("dispAllKey")
    .taskControlKey.Value = setVal("taskControlKey")
    .ScaleKey.Value = setVal("ScaleKey")
    
    '�S����
    .Assign01.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 4)
    .Assign02.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 5)
    .Assign03.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 6)
    .Assign04.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 7)
    .Assign05.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 8)
    .Assign06.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 9)
    .Assign07.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 10)
    .Assign08.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 11)
    .Assign09.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 12)
    .Assign10.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 13)
    .Assign11.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 14)
    .Assign12.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 15)
    .Assign13.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 16)
    .Assign14.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 17)
    .Assign15.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 18)
    .Assign16.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 19)
    .Assign17.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 20)
    .Assign18.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 21)
    .Assign19.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 22)
    .Assign20.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 23)
    .Assign21.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 24)
    .Assign22.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 25)
    .Assign23.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 26)
    .Assign24.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 27)
    .Assign25.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 28)
    .Assign26.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 29)
    .Assign27.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 30)
    .Assign28.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 31)
    .Assign29.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 32)
    .Assign30.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 33)
    .Assign31.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 34)
    .Assign32.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 35)
    .Assign33.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 36)
    .Assign34.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 37)
    .Assign35.Text = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 38)
    
    '�S���ҐF
    .AssignColor01.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 4).Interior.Color
    .AssignColor02.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 5).Interior.Color
    .AssignColor03.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 6).Interior.Color
    .AssignColor04.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 7).Interior.Color
    .AssignColor05.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 8).Interior.Color
    .AssignColor06.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 9).Interior.Color
    .AssignColor07.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 10).Interior.Color
    .AssignColor08.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 11).Interior.Color
    .AssignColor09.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 12).Interior.Color
    .AssignColor10.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 13).Interior.Color
    .AssignColor11.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 14).Interior.Color
    .AssignColor12.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 15).Interior.Color
    .AssignColor13.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 16).Interior.Color
    .AssignColor14.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 17).Interior.Color
    .AssignColor15.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 18).Interior.Color
    .AssignColor16.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 19).Interior.Color
    .AssignColor17.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 20).Interior.Color
    .AssignColor18.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 21).Interior.Color
    .AssignColor19.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 22).Interior.Color
    .AssignColor20.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 23).Interior.Color
    .AssignColor21.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 24).Interior.Color
    .AssignColor22.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 25).Interior.Color
    .AssignColor23.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 26).Interior.Color
    .AssignColor24.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 27).Interior.Color
    .AssignColor25.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 28).Interior.Color
    .AssignColor26.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 29).Interior.Color
    .AssignColor27.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 30).Interior.Color
    .AssignColor28.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 31).Interior.Color
    .AssignColor29.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 32).Interior.Color
    .AssignColor30.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 33).Interior.Color
    .AssignColor31.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 34).Interior.Color
    .AssignColor32.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 35).Interior.Color
    .AssignColor33.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 36).Interior.Color
    .AssignColor34.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 37).Interior.Color
    .AssignColor35.BackColor = Sh_wbsConfig.Range(setVal("cell_AssignorList") & 38).Interior.Color
    
    '�S���ҒP��
    .unitCost01.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 4)
    .unitCost02.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 5)
    .unitCost03.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 6)
    .unitCost04.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 7)
    .unitCost05.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 8)
    .unitCost06.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 9)
    .unitCost07.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 10)
    .unitCost08.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 11)
    .unitCost09.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 12)
    .unitCost10.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 13)
    .unitCost11.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 14)
    .unitCost12.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 15)
    .unitCost13.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 16)
    .unitCost14.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 17)
    .unitCost15.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 18)
    .unitCost16.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 19)
    .unitCost17.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 20)
    .unitCost18.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 21)
    .unitCost19.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 22)
    .unitCost20.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 23)
    .unitCost21.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 24)
    .unitCost22.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 25)
    .unitCost23.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 26)
    .unitCost24.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 27)
    .unitCost25.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 28)
    .unitCost26.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 29)
    .unitCost27.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 30)
    .unitCost28.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 31)
    .unitCost29.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 32)
    .unitCost30.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 33)
    .unitCost31.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 34)
    .unitCost32.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 35)
    .unitCost33.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 36)
    .unitCost34.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 37)
    .unitCost35.Text = Sh_wbsConfig.Range(setVal("cell_unitCostorList") & 38)

    
    '��Ўw��x��
    For line = 3 To Sh_wbsConfig.Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).Row
      If Sh_wbsConfig.Range(setVal("cell_CompanyHoliday") & line) <> "" Then
        If CompanyHolidayList = "" Then
          CompanyHolidayList = Sh_wbsConfig.Range(setVal("cell_CompanyHoliday") & line)
        Else
          CompanyHolidayList = CompanyHolidayList & vbNewLine & Sh_wbsConfig.Range(setVal("cell_CompanyHoliday") & line)
        End If
      End If
    Next
    .CompanyHoliday.Text = CompanyHolidayList
    
    '���o�^�X�N
    For line = 3 To Sh_wbsConfig.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).Row
      If Sh_wbsConfig.Range(setVal("cell_DataExtract") & line) <> "" Then
        If dataExtractList = "" Then
          dataExtractList = Sh_wbsConfig.Range(setVal("cell_DataExtract") & line)
        Else
          dataExtractList = dataExtractList & vbNewLine & Sh_wbsConfig.Range(setVal("cell_DataExtract") & line)
        End If
      End If
    Next
    .dataExtract.Text = dataExtractList
    
    
    '�\���ݒ�
    .view_Plan.Value = setVal("view_Plan")
    .view_Assign.Value = setVal("view_Assign")
    .view_Progress.Value = setVal("view_Progress")
    .view_Achievement.Value = setVal("view_Achievement")
    .view_Task.Value = setVal("view_Task")
    .view_TaskInfo.Value = setVal("view_TaskInfo")
    .view_TaskAllocation.Value = setVal("view_TaskAllocation")
    .view_LineInfo.Value = setVal("view_LineInfo")
    
    .view_WorkLoad.Value = setVal("view_WorkLoad")
    .view_LateOrEarly.Value = setVal("view_LateOrEarly")
    .view_Note.Value = setVal("view_Note")
    
    .viewGant_TaskName.Value = setVal("viewGant_TaskName")
    .viewGant_Assignor.Value = setVal("viewGant_Assignor")
  
  End With
  
  Frm_Option.Show

  Exit Function
'�G���[������------------------------------------

catchError:

End Function


'==================================================================================================
Function �I�v�V�����ݒ�l�i�[()

  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim CompanyHoliday As Variant, Holiday As Variant, dataExtract As Variant

  On Error Resume Next
  
  Call Ctl_ProgressBar.showStart
  Sh_wbsConfig.Select
'  For line = 3 To Sh_wbsConfig.Range("B5")
'    'Call Ctl_ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, Sh_wbsConfig.Range("B5"), Sh_wbsConfig.Range("A" & line) & ":" & getVal(Sh_wbsConfig.Range("A" & line)))
'    Sh_wbsConfig.Range(Sh_wbsConfig.Range("A" & line)).Select
'
'    If IsEmpty(getVal(Sh_wbsConfig.Range("A" & line))) = False Then
'      Select Case Sh_wbsConfig.Range("A" & line)
'        Case "baseDay"
'          If getVal(Sh_wbsConfig.Range("A" & line)) = Format(Now, "yyyy/mm/dd") Then
'            Sh_wbsConfig.Range(Sh_wbsConfig.Range("A" & line)).FormulaR1C1 = "=TODAY()"
'          Else
'            Sh_wbsConfig.Range(Sh_wbsConfig.Range("A" & line)) = getVal(Sh_wbsConfig.Range("A" & line))
'          End If
'
'        Case ""
'        Case Else
'          Sh_wbsConfig.Range(Sh_wbsConfig.Range("A" & line)) = getVal(Sh_wbsConfig.Range("A" & line))
'      End Select
'    End If
'  Next
  
'  '�V���[�g�J�b�g�L�[�̐ݒ�
'  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
'  For line = 3 To endLine
'    Call Ctl_ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, Sh_wbsConfig.Range("B5"), "�V���[�g�J�b�g�L�[�ݒ�")
'
'    Range(Range(setVal("cell_ShortcutFuncName") & line)).Select
'    Range(Range(setVal("cell_ShortcutFuncName") & line)) = getVal(Range(setVal("cell_ShortcutFuncName") & line))
'  Next
'
  '��Ўw��x���̐ݒ�
  line = 3
  Sh_wbsConfig.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).Row).ClearContents
  For Each CompanyHoliday In Split(getVal("CompanyHoliday"), vbNewLine)
    DoEvents
    Sh_wbsConfig.Range(setVal("cell_CompanyHoliday") & line) = CompanyHoliday
    line = line + 1
  Next
  

  
  
  
'
'  '���o�^�X�N�̐ݒ�
'  line = 3
'  Sh_wbsConfig.Range(setVal("cell_DataExtract") & "3:" & setVal("cell_DataExtract") & Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row).ClearContents
'  For Each dataExtract In Split(getVal("dataExtract"), vbNewLine)
'    Call Ctl_ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, 100, "���o�^�X�N�̐ݒ�")
'
'    Sh_wbsConfig.Range(setVal("cell_DataExtract") & line) = dataExtract
'    line = line + 1
'  Next


  '�S����
  Sh_wbsConfig.Range(setVal("cell_AssignorList") & "4:" & setVal("cell_AssignorList") & Cells(Rows.count, Library.getColumnNo(setVal("cell_AssignorList"))).End(xlUp).Row).Clear
  For line = 4 To 38
    'Call Ctl_ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, 38, "�S����:" & getVal("Assign" & Format(line - 3, "00")))
    
    Sh_wbsConfig.Range(setVal("cell_AssignorList") & line) = getVal("Assign" & Format(line - 3, "00"))
    Sh_wbsConfig.Range(setVal("cell_AssignorList") & line).Interior.Color = getVal("AssignColor" & Format(line - 3, "00"))
    
    Sh_wbsConfig.Range(setVal("cell_unitCostorList") & line) = getVal("unitCost" & Format(line - 3, "00"))
    
  Next
  Sh_wbsConfig.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & 38).Select

'  Call �r��.�͂݌r��
'  Call Menu.M_�V���[�g�J�b�g�ݒ�
  
  Call WBS_Option.�ݒ�V�[�g�R�s�[("forWBS")
  ActiveWorkbook.Worksheets("WBS").Select
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  
  '�S���҃��X�g����------------------------------
  Call Ctl_Assign.�S���҃��X�g�\��

  endLine = Range("A1").SpecialCells(xlLastCell).Row
  With Range("K" & startLine & ":L" & endLine).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(lstAssign, ",")
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOn
    .ShowInput = True
    .showError = False
  End With
  
  
  
  
  

  Call Ctl_ProgressBar.showEnd
  Set getVal = Nothing
  
End Function