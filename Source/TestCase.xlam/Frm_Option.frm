VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "オプション - WBS"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   OleObjectBlob   =   "Frm_Option.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************************************************
'   カレンダーフォーム4(日付入力部品)   ※テスト用フォーム      UserForm1(UserForm)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'変更日付 Rev  変更履歴内容========================================================================>
'18/02/21(1.00)新規作成
'18/11/28(1.80)カレンダーフォーム名、モジュール名変更
'***************************************************************************************************
Option Explicit
'===================================================================================================
Private Const g_cnsAddLeft As Long = 3                          ' Left調整値
Private Const g_cnsAddTop As Long = 19                          ' Top調整値
Private Const g_cnsAddLeft2 As Long = 4                         ' Left調整値(フレーム用)
Private Const g_cnsAddTop2 As Long = 25                         ' Top調整値(フレーム用)
' ※これらの調整値はWindows10時点の画面で適当に見繕った値です。
' 　二重にフレームが重なった等の場合は別途調整が必要です。


#If Win64 Then
  Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
  Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
  Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vkey As Long) As Integer
  Private Declare Function GetForegroundWindow Lib "user32" () As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If


Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&

Public KeyPressFlg As Boolean



'Private Sub UserForm_Activate()
'    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'    Me.StartUpPosition = 1
'End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Initialize
'* 機能　　：ユーザーフォームの初期化
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2018年02月21日
'* 作成者　：井上　治
'* 更新日　：2018年02月21日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_Initialize()
  
  'マウスカーソルを標準に設定
  Application.Cursor = xlDefault

  ' テキストボックスに▼ボタンを表示させる
  startDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  endDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
  baseDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
End Sub


'***************************************************************************************************
' ■■■ フォームイベント ■■■
'***************************************************************************************************
'* 処理名　：startDay_DropButtonClick
'* 機能　　：フォーム上のテキストボックスイベント(DropButtonClick)
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2018年02月21日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub startDay_DropButtonClick()
    '==============================================================================================-
    Dim lngLeft As Long                                             ' 位置(横方向)
    Dim lngTop As Long                                              ' 位置(縦方向)
    ' フォーム+テキストボックスのLeft,Top値から位置を判定
    lngLeft = Me.Left + startDay.Left + g_cnsAddLeft
    lngTop = Me.Top + startDay.Top + startDay.Height + g_cnsAddTop
    '==============================================================================================-
    ' カレンダーフォームを起動する
    Call modCalendar5R.ShowCalendarFromTextBox2(startDay, lngLeft, lngTop, "開始日選択")
End Sub

Private Sub endDay_DropButtonClick()
    '==============================================================================================-
    Dim lngLeft As Long                                             ' 位置(横方向)
    Dim lngTop As Long                                              ' 位置(縦方向)
    ' フォーム+テキストボックスのLeft,Top値から位置を判定
    lngLeft = Me.Left + endDay.Left + g_cnsAddLeft
    lngTop = Me.Top + endDay.Top + endDay.Height + g_cnsAddTop
    '==============================================================================================-
    ' カレンダーフォームを起動する
    Call modCalendar5R.ShowCalendarFromTextBox2(endDay, lngLeft, lngTop, "終了日選択")
End Sub

Private Sub baseDay_DropButtonClick()
    '==============================================================================================-
    Dim lngLeft As Long                                             ' 位置(横方向)
    Dim lngTop As Long                                              ' 位置(縦方向)
    ' フォーム+テキストボックスのLeft,Top値から位置を判定
    lngLeft = Me.Left + baseDay.Left + g_cnsAddLeft
    lngTop = Me.Top + baseDay.Top + baseDay.Height + g_cnsAddTop
    '==============================================================================================-
    ' カレンダーフォームを起動する
    Call modCalendar5R.ShowCalendarFromTextBox2(baseDay, lngLeft, lngTop, "基準日選択")
End Sub


'***************************************************************************************************
'* 処理名　：GP_GakuEnter
'* 機能　　：金額項目入力用編集
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = テキストボックス(Object)
'===================================================================================================
'* 作成日　：2003年07月25日
'* 作成者　：井上　治
'* 更新日　：2020年02月24日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_GakuEnter(objTextBox As MSForms.TextBox)
    '==============================================================================================-
    Dim strGaku As String                                           ' 入力テキスト
    Dim crnGaku As Currency                                         ' 金額値
    strGaku = Trim(objTextBox.Text)
    ' 数値か
    If IsNumeric(strGaku) Then
        crnGaku = CCur(strGaku)
        ' 3桁カンマ除去で編集
        objTextBox.Text = Format(crnGaku, "0")
        ' 全桁選択
        Call GP_AllSelect(objTextBox)
    End If
End Sub
'***************************************************************************************************
'* 処理名　：FP_GakuExit
'* 機能　　：金額項目表示用編集
'===================================================================================================
'* 返り値　：チェック成否(Boolean)
'* 引数　　：Arg1 = テキストボックス(Object)
'===================================================================================================
'* 作成日　：2003年07月25日
'* 作成者　：井上　治
'* 更新日　：2020年02月24日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Function FP_GakuExit(objTextBox As MSForms.TextBox) As Boolean
    '==============================================================================================-
    Dim strGaku As String                                           ' 入力テキスト
    Dim crnGaku As Currency                                         ' 金額値
    FP_GakuExit = False
    strGaku = Trim(objTextBox.Text)
    ' 数値か
    If IsNumeric(strGaku) Then
        crnGaku = CCur(strGaku)
        ' 3桁カンマ付きで編集
        objTextBox.Text = Format(crnGaku, "#,##0")
        FP_GakuExit = True
    ElseIf strGaku = "" Then
    Else
        MsgBox "数字ではありません。", vbExclamation
        ' 全桁選択
        Call GP_AllSelect(objTextBox)
    End If
End Function
'***************************************************************************************************
'* 処理名　：GP_AllSelect
'* 機能　　：全桁選択
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = テキストボックス(Object)
'===================================================================================================
'* 作成日　：2003年07月25日
'* 作成者　：井上　治
'* 更新日　：2020年02月24日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_AllSelect(objTextBox As MSForms.TextBox)
    '==============================================================================================-
    With objTextBox
      .SetFocus
      .SelStart = 0
      .SelLength = Len(.Text)
    End With
End Sub

Function chkScope(minVal As MSForms.TextBox, maxVal As MSForms.TextBox)

  If minVal.Text <= maxVal.Text Then
    chkScope = True
  ElseIf maxVal.Text = 0 Then
    chkScope = True
  Else
    message.Caption = "範囲設定が成立しません"
    chkScope = False
  End If
  
End Function

'**************************************************************************************************
' * 担当者情報のクリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub clearAssignor_Click()
    Assign01.Text = ""
    Assign02.Text = ""
    Assign03.Text = ""
    Assign04.Text = ""
    Assign05.Text = ""
    Assign06.Text = ""
    Assign07.Text = ""
    Assign08.Text = ""
    Assign09.Text = ""
    Assign10.Text = ""
    Assign11.Text = ""
    Assign12.Text = ""
    Assign13.Text = ""
    Assign14.Text = ""
    Assign15.Text = ""
    Assign16.Text = ""
    Assign17.Text = ""
    Assign18.Text = ""
    Assign19.Text = ""
    Assign20.Text = ""
    Assign21.Text = ""
    Assign22.Text = ""
    Assign23.Text = ""
    Assign24.Text = ""
'    Assign25.Text = ""
'    Assign26.Text = ""
'    Assign27.Text = ""
'    Assign28.Text = ""
'    Assign29.Text = ""
'    Assign30.Text = ""
'    Assign31.Text = ""
'    Assign32.Text = ""
'    Assign33.Text = ""
'    Assign34.Text = ""
'    Assign35.Text = ""
    
    
    AssignColor01.BackColor = 16777215
    AssignColor02.BackColor = 16777215
    AssignColor03.BackColor = 16777215
    AssignColor04.BackColor = 16777215
    AssignColor05.BackColor = 16777215
    AssignColor06.BackColor = 16777215
    AssignColor07.BackColor = 16777215
    AssignColor08.BackColor = 16777215
    AssignColor09.BackColor = 16777215
    AssignColor10.BackColor = 16777215
    AssignColor11.BackColor = 16777215
    AssignColor12.BackColor = 16777215
    AssignColor13.BackColor = 16777215
    AssignColor14.BackColor = 16777215
    AssignColor15.BackColor = 16777215
    AssignColor16.BackColor = 16777215
    AssignColor17.BackColor = 16777215
    AssignColor18.BackColor = 16777215
    AssignColor19.BackColor = 16777215
    AssignColor20.BackColor = 16777215
    AssignColor21.BackColor = 16777215
    AssignColor22.BackColor = 16777215
    AssignColor23.BackColor = 16777215
    AssignColor24.BackColor = 16777215
'    AssignColor25.BackColor = 16777215
'    AssignColor26.BackColor = 16777215
'    AssignColor27.BackColor = 16777215
'    AssignColor28.BackColor = 16777215
'    AssignColor29.BackColor = 16777215
'    AssignColor30.BackColor = 16777215
'    AssignColor31.BackColor = 16777215
'    AssignColor32.BackColor = 16777215
'    AssignColor33.BackColor = 16777215
'    AssignColor34.BackColor = 16777215
'    AssignColor35.BackColor = 16777215



End Sub


'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  Set getVal = New Collection
  
  With getVal
    .Add item:=startDay.Text, key:="startDay"
    .Add item:=endDay.Text, key:="endDay"
    .Add item:=baseDay.Text, key:="baseDay"
    .Add item:=setLightning.value, key:="setLightning"
    .Add item:=setDispProgress100.value, key:="setDispProgress100"
    .Add item:=CompanyHoliday.value, key:="CompanyHoliday"
    .Add item:=dataExtract.value, key:="dataExtract"
    
    'スタイル関連
    .Add item:=lineColor.BackColor, key:="lineColor"
    
    .Add item:=SaturdayColor.BackColor, key:="SaturdayColor"
    .Add item:=SundayColor.BackColor, key:="SundayColor"
    .Add item:=CompanyHolidayColor.BackColor, key:="CompanyHolidayColor"

    .Add item:=lineColor_Plan.BackColor, key:="lineColor_Plan"
    .Add item:=lineColor_Achievement.BackColor, key:="lineColor_Achievement"
    .Add item:=lineColor_Lightning.BackColor, key:="lineColor_Lightning"
    .Add item:=lineColor_TaskLevel1.BackColor, key:="lineColor_TaskLevel1"
    .Add item:=lineColor_TaskLevel2.BackColor, key:="lineColor_TaskLevel2"
    .Add item:=lineColor_TaskLevel3.BackColor, key:="lineColor_TaskLevel3"

    'ショートカットキー設定
    .Add item:=optionKey.Text, key:="optionKey"
    .Add item:=centerKey.Text, key:="centerKey"
    .Add item:=filterKey.Text, key:="filterKey"
    .Add item:=clearFilterKey.Text, key:="clearFilterKey"
    .Add item:=taskCheckKey.Text, key:="taskCheckKey"
    .Add item:=makeGanttKey.Text, key:="makeGanttKey"
    .Add item:=clearGanttKey.Text, key:="clearGanttKey"
    .Add item:=dispAllKey.Text, key:="dispAllKey"
    .Add item:=taskControlKey.Text, key:="taskControlKey"
    .Add item:=ScaleKey.Text, key:="ScaleKey"

    '担当者
    .Add item:=Assign01.Text, key:="Assign01"
    .Add item:=Assign02.Text, key:="Assign02"
    .Add item:=Assign03.Text, key:="Assign03"
    .Add item:=Assign04.Text, key:="Assign04"
    .Add item:=Assign05.Text, key:="Assign05"
    .Add item:=Assign06.Text, key:="Assign06"
    .Add item:=Assign07.Text, key:="Assign07"
    .Add item:=Assign08.Text, key:="Assign08"
    .Add item:=Assign09.Text, key:="Assign09"
    .Add item:=Assign10.Text, key:="Assign10"
    .Add item:=Assign11.Text, key:="Assign11"
    .Add item:=Assign12.Text, key:="Assign12"
    .Add item:=Assign13.Text, key:="Assign13"
    .Add item:=Assign14.Text, key:="Assign14"
    .Add item:=Assign15.Text, key:="Assign15"
    .Add item:=Assign16.Text, key:="Assign16"
    .Add item:=Assign17.Text, key:="Assign17"
    .Add item:=Assign18.Text, key:="Assign18"
    .Add item:=Assign19.Text, key:="Assign19"
    .Add item:=Assign20.Text, key:="Assign20"
    .Add item:=Assign21.Text, key:="Assign21"
    .Add item:=Assign22.Text, key:="Assign22"
    .Add item:=Assign23.Text, key:="Assign23"
    .Add item:=Assign24.Text, key:="Assign24"
    .Add item:=Assign25.Text, key:="Assign25"
    .Add item:=Assign26.Text, key:="Assign26"
    .Add item:=Assign27.Text, key:="Assign27"
    .Add item:=Assign28.Text, key:="Assign28"
    .Add item:=Assign29.Text, key:="Assign29"
    .Add item:=Assign30.Text, key:="Assign30"
    .Add item:=Assign31.Text, key:="Assign31"
    .Add item:=Assign32.Text, key:="Assign32"
    .Add item:=Assign33.Text, key:="Assign33"
    .Add item:=Assign34.Text, key:="Assign34"
    .Add item:=Assign35.Text, key:="Assign35"

    '担当者Color
    .Add item:=AssignColor01.BackColor, key:="AssignColor01"
    .Add item:=AssignColor02.BackColor, key:="AssignColor02"
    .Add item:=AssignColor03.BackColor, key:="AssignColor03"
    .Add item:=AssignColor04.BackColor, key:="AssignColor04"
    .Add item:=AssignColor05.BackColor, key:="AssignColor05"
    .Add item:=AssignColor06.BackColor, key:="AssignColor06"
    .Add item:=AssignColor07.BackColor, key:="AssignColor07"
    .Add item:=AssignColor08.BackColor, key:="AssignColor08"
    .Add item:=AssignColor09.BackColor, key:="AssignColor09"
    .Add item:=AssignColor10.BackColor, key:="AssignColor10"
    .Add item:=AssignColor11.BackColor, key:="AssignColor11"
    .Add item:=AssignColor12.BackColor, key:="AssignColor12"
    .Add item:=AssignColor13.BackColor, key:="AssignColor13"
    .Add item:=AssignColor14.BackColor, key:="AssignColor14"
    .Add item:=AssignColor15.BackColor, key:="AssignColor15"
    .Add item:=AssignColor16.BackColor, key:="AssignColor16"
    .Add item:=AssignColor17.BackColor, key:="AssignColor17"
    .Add item:=AssignColor18.BackColor, key:="AssignColor18"
    .Add item:=AssignColor19.BackColor, key:="AssignColor19"
    .Add item:=AssignColor20.BackColor, key:="AssignColor20"
    .Add item:=AssignColor21.BackColor, key:="AssignColor21"
    .Add item:=AssignColor22.BackColor, key:="AssignColor22"
    .Add item:=AssignColor23.BackColor, key:="AssignColor23"
    .Add item:=AssignColor24.BackColor, key:="AssignColor24"
    .Add item:=AssignColor25.BackColor, key:="AssignColor25"
    .Add item:=AssignColor26.BackColor, key:="AssignColor26"
    .Add item:=AssignColor27.BackColor, key:="AssignColor27"
    .Add item:=AssignColor28.BackColor, key:="AssignColor28"
    .Add item:=AssignColor29.BackColor, key:="AssignColor29"
    .Add item:=AssignColor30.BackColor, key:="AssignColor30"
    .Add item:=AssignColor31.BackColor, key:="AssignColor31"
    .Add item:=AssignColor32.BackColor, key:="AssignColor32"
    .Add item:=AssignColor33.BackColor, key:="AssignColor33"
    .Add item:=AssignColor34.BackColor, key:="AssignColor34"
    .Add item:=AssignColor35.BackColor, key:="AssignColor35"
    
    '担当者単価
    .Add item:=unitCost01.Text, key:="unitCost01"
    .Add item:=unitCost02.Text, key:="unitCost02"
    .Add item:=unitCost03.Text, key:="unitCost03"
    .Add item:=unitCost04.Text, key:="unitCost04"
    .Add item:=unitCost05.Text, key:="unitCost05"
    .Add item:=unitCost06.Text, key:="unitCost06"
    .Add item:=unitCost07.Text, key:="unitCost07"
    .Add item:=unitCost08.Text, key:="unitCost08"
    .Add item:=unitCost09.Text, key:="unitCost09"
    .Add item:=unitCost10.Text, key:="unitCost10"
    .Add item:=unitCost11.Text, key:="unitCost11"
    .Add item:=unitCost12.Text, key:="unitCost12"
    .Add item:=unitCost13.Text, key:="unitCost13"
    .Add item:=unitCost14.Text, key:="unitCost14"
    .Add item:=unitCost15.Text, key:="unitCost15"
    .Add item:=unitCost16.Text, key:="unitCost16"
    .Add item:=unitCost17.Text, key:="unitCost17"
    .Add item:=unitCost18.Text, key:="unitCost18"
    .Add item:=unitCost19.Text, key:="unitCost19"
    .Add item:=unitCost20.Text, key:="unitCost20"
    .Add item:=unitCost21.Text, key:="unitCost21"
    .Add item:=unitCost22.Text, key:="unitCost22"
    .Add item:=unitCost23.Text, key:="unitCost23"
    .Add item:=unitCost24.Text, key:="unitCost24"
    .Add item:=unitCost25.Text, key:="unitCost25"
    .Add item:=unitCost26.Text, key:="unitCost26"
    .Add item:=unitCost27.Text, key:="unitCost27"
    .Add item:=unitCost28.Text, key:="unitCost28"
    .Add item:=unitCost29.Text, key:="unitCost29"
    .Add item:=unitCost30.Text, key:="unitCost30"
    .Add item:=unitCost31.Text, key:="unitCost31"
    .Add item:=unitCost32.Text, key:="unitCost32"
    .Add item:=unitCost33.Text, key:="unitCost33"
    .Add item:=unitCost34.Text, key:="unitCost34"
    .Add item:=unitCost35.Text, key:="unitCost35"
    
    
    '表示設定
    .Add item:=view_Plan.value, key:="view_Plan"
    .Add item:=view_Assign.value, key:="view_Assign"
    .Add item:=view_Progress.value, key:="view_Progress"
    .Add item:=view_Achievement.value, key:="view_Achievement"
    .Add item:=view_Task.value, key:="view_Task"
    
    .Add item:=view_WorkLoad.value, key:="view_WorkLoad"
    .Add item:=view_LateOrEarly.value, key:="view_LateOrEarly"
    .Add item:=view_Note.value, key:="view_Note"

    .Add item:=view_TaskInfo.value, key:="view_TaskInfo"
    .Add item:=view_TaskAllocation.value, key:="view_TaskAllocation"
    .Add item:=view_LineInfo.value, key:="view_LineInfo"

    .Add item:=viewGant_TaskName.value, key:="viewGant_TaskName"
    .Add item:=viewGant_Assignor.value, key:="viewGant_Assignor"

  End With
  Unload Me
  
  
  Call Ctl_Option.オプション設定値格納

End Sub


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()
  Unload Me
  Call Library.endScript
  End
End Sub

'**************************************************************************************************
' * スタイル関連
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub lineColor_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.lineColor.BackColor)
  Me.lineColor.BackColor = colorValue
End Sub
Private Sub SaturdayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.SaturdayColor.BackColor)
  Me.SaturdayColor.BackColor = colorValue
End Sub
Private Sub SundayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.SundayColor.BackColor)
  Me.SundayColor.BackColor = colorValue
End Sub
Private Sub CompanyHolidayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.CompanyHolidayColor.BackColor)
  Me.CompanyHolidayColor.BackColor = colorValue
End Sub
Private Sub lineColor_Plan_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Plan.BackColor)
  Me.lineColor_Plan.BackColor = colorValue
End Sub
Private Sub lineColor_Achievement_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Achievement.BackColor)
  Me.lineColor_Achievement.BackColor = colorValue
End Sub
Private Sub lineColor_Lightning_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Lightning.BackColor)
  Me.lineColor_Lightning.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel1_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel1.BackColor)
  Me.lineColor_TaskLevel1.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel2_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel2.BackColor)
  Me.lineColor_TaskLevel2.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel3_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel3.BackColor)
  Me.lineColor_TaskLevel3.BackColor = colorValue
End Sub


'**************************************************************************************************
' * 担当者
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

Private Sub AssignColor01_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor01.BackColor)
  Me.AssignColor01.BackColor = colorValue
End Sub

Private Sub AssignColor02_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor02.BackColor)
  Me.AssignColor02.BackColor = colorValue
End Sub

Private Sub AssignColor03_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor03.BackColor)
  Me.AssignColor03.BackColor = colorValue
End Sub

Private Sub AssignColor04_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor04.BackColor)
  Me.AssignColor04.BackColor = colorValue
End Sub

Private Sub AssignColor05_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor05.BackColor)
  Me.AssignColor05.BackColor = colorValue
End Sub

Private Sub AssignColor06_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor06.BackColor)
  Me.AssignColor06.BackColor = colorValue
End Sub

Private Sub AssignColor07_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor07.BackColor)
  Me.AssignColor07.BackColor = colorValue
End Sub

Private Sub AssignColor08_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor08.BackColor)
  Me.AssignColor08.BackColor = colorValue
End Sub

Private Sub AssignColor09_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor09.BackColor)
  Me.AssignColor09.BackColor = colorValue
End Sub

Private Sub AssignColor10_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor10.BackColor)
  Me.AssignColor10.BackColor = colorValue
End Sub

Private Sub AssignColor11_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor11.BackColor)
  Me.AssignColor11.BackColor = colorValue
End Sub

Private Sub AssignColor12_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor12.BackColor)
  Me.AssignColor12.BackColor = colorValue
End Sub

Private Sub AssignColor13_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor13.BackColor)
  Me.AssignColor13.BackColor = colorValue
End Sub

Private Sub AssignColor14_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor14.BackColor)
  Me.AssignColor14.BackColor = colorValue
End Sub

Private Sub AssignColor15_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor15.BackColor)
  Me.AssignColor15.BackColor = colorValue
End Sub

Private Sub AssignColor16_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor16.BackColor)
  Me.AssignColor16.BackColor = colorValue
End Sub

Private Sub AssignColor17_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor17.BackColor)
  Me.AssignColor17.BackColor = colorValue
End Sub

Private Sub AssignColor18_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor18.BackColor)
  Me.AssignColor18.BackColor = colorValue
End Sub

Private Sub AssignColor19_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor19.BackColor)
  Me.AssignColor19.BackColor = colorValue
End Sub

Private Sub AssignColor20_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor20.BackColor)
  Me.AssignColor20.BackColor = colorValue
End Sub

Private Sub AssignColor21_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor21.BackColor)
  Me.AssignColor21.BackColor = colorValue
End Sub

Private Sub AssignColor22_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor22.BackColor)
  Me.AssignColor22.BackColor = colorValue
End Sub

Private Sub AssignColor23_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor23.BackColor)
  Me.AssignColor23.BackColor = colorValue
End Sub

Private Sub AssignColor24_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor24.BackColor)
  Me.AssignColor24.BackColor = colorValue
End Sub

Private Sub AssignColor25_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor25.BackColor)
  Me.AssignColor25.BackColor = colorValue
End Sub

Private Sub AssignColor26_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor26.BackColor)
  Me.AssignColor26.BackColor = colorValue
End Sub

Private Sub AssignColor27_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor27.BackColor)
  Me.AssignColor27.BackColor = colorValue
End Sub

Private Sub AssignColor28_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor28.BackColor)
  Me.AssignColor28.BackColor = colorValue
End Sub

Private Sub AssignColor29_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor29.BackColor)
  Me.AssignColor29.BackColor = colorValue
End Sub

Private Sub AssignColor30_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor30.BackColor)
  Me.AssignColor30.BackColor = colorValue
End Sub

Private Sub AssignColor31_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor31.BackColor)
  Me.AssignColor31.BackColor = colorValue
End Sub

Private Sub AssignColor32_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor32.BackColor)
  Me.AssignColor32.BackColor = colorValue
End Sub

Private Sub AssignColor33_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor33.BackColor)
  Me.AssignColor33.BackColor = colorValue
End Sub

Private Sub AssignColor34_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor34.BackColor)
  Me.AssignColor34.BackColor = colorValue
End Sub

Private Sub AssignColor35_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor35.BackColor)
  Me.AssignColor35.BackColor = colorValue
End Sub

