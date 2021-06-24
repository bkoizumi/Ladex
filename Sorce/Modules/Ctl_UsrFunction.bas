Attribute VB_Name = "Ctl_UsrFunction"
'**************************************************************************************************
' * ユーザー定義関数
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function InitializeUsrFunction()

  Application.MacroOptions _
    Macro:="chkWorkDay", _
    Description:="第N営業日かチェックし、True/Falseを返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする日付を指定", "第N営業日を数値で指定")

  Application.MacroOptions _
    Macro:="getWorkDay", _
    Description:="第N営業日をシリアル値で返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする年を数値で指定", "チェックする月を数値で指定", "第N営業日を数値で指定")

  Application.MacroOptions _
    Macro:="chkWeekNum", _
    Description:="第N週X曜日の日付をシリアル値で返す", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("チェックする日付を指定", "第N週を数値で指定", "X曜日を数値で指定")

  Application.MacroOptions _
    Macro:="Textjoin", _
    Description:="文字列連結", _
    Category:=thisAppName, _
    ArgumentDescriptions:=Array("区切り文字", "空欄時処理[True：処理する/False：処理しない]", "文字列1,文字列2, ...")


End Function



'==================================================================================================
Function Textjoin(Delim, Ignore As Boolean, ParamArray par())
Attribute Textjoin.VB_Description = "文字列連結"
Attribute Textjoin.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim i As Integer
  Dim tR As Range

  Textjoin = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Or Ignore = False Then
          Textjoin = Textjoin & Delim & tR.Value2
        End If
      Next
    Else
      If (par(i) <> "" And par(i) <> "<>") Or Ignore = False Then
        Textjoin = Textjoin & Delim & par(i)
      End If
    End If
  Next

  Textjoin = Mid(Textjoin, Len(Delim) + 1)

End Function




'==================================================================================================
Function chkWorkDay(ByVal checkDate As Date, ByVal bsnDay As Long) As Boolean
Attribute chkWorkDay.VB_Description = "第N営業日かチェックし、True/Falseを返す"
Attribute chkWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
  
  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(Year(checkDate), Month(checkDate), 1)
  getDay = Application.WorksheetFunction.WorkDay(firstDay, bsnDay, arryHollyday)
  
  If checkDate = getDay Then
    chkWorkDay = True
  Else
    chkWorkDay = False
  End If

End Function

'==================================================================================================
Function chkWeekNum(ByVal checkDate As Date, ByVal checkWeekday As Long, ByVal weekNum As Long) As Date
Attribute chkWeekNum.VB_Description = "第N週X曜日の日付をシリアル値で返す"
Attribute chkWeekNum.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Long, diff As Long
  
  Application.Volatile
  
  firstDay = Weekday(DateSerial(Year(checkDate), Month(checkDate), 1))
  diff = (checkWeekday + 7 - firstDay) Mod 7
  getDay = DateSerial(Year(checkDate), Month(checkDate), 1 + diff + 7 * (weekNum - 1))
  
  chkWeekNum = getDay
  
End Function

'==================================================================================================
Function getWorkDay(ByVal cYear As Long, ByVal cMonth As Long, ByVal bsnDay As Long) As Date
Attribute getWorkDay.VB_Description = "第N営業日をシリアル値で返す"
Attribute getWorkDay.VB_ProcData.VB_Invoke_Func = " \n19"
  Dim getDay As Date, firstDay As Date
  
  Application.Volatile
  If Library.chkArrayEmpty(arryHollyday) = True Then
    Call Ctl_Hollyday.InitializeHollyday
  End If
  
  firstDay = DateSerial(cYear, cMonth, 1)
  getWorkDay = Application.WorksheetFunction.WorkDay(firstDay - 1, bsnDay, arryHollyday)
  
End Function

