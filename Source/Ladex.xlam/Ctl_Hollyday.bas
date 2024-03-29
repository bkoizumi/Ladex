Attribute VB_Name = "Ctl_Hollyday"
Option Explicit

'******************************************************************************
' ある日が祝日であるか？その場合どの祝日か？を調べる関数。
' http://www.excelio.jp/LABORATORY/EXCEL_CALENDER.html
'******************************************************************************

'==================================================================================================
Function InitializeHollyday()
  Dim startDay As Date, endDay As Date
  Dim NowYear As Integer, today As Date
  Dim HollydayName As String
  Dim count As Long
  Const funcName As String = "Ctl_Hollyday.InitializeHollyday"
  
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
  
  NowYear = Format(Date, "yyyy")
  startDay = Format(NowYear & "/4/1", "yyyy/mm/dd")
  endDay = Format(NowYear + 2 & "/3/31", "yyyy/mm/dd")
  
  If Library.chkArrayEmpty(arryHollyday) = False Then
    GoTo Lbl_exitFunction
  End If
  
  count = 0
  For today = startDay To endDay
    If GetHollyday(today, HollydayName) = True Then
      ReDim Preserve arryHollyday(count)
      arryHollyday(count) = today
    End If
  Next

Lbl_exitFunction:
  '処理終了--------------------------------------
  If runFlg = False Then
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
  Call Library.showDebugForm(funcName, " [" & Err.Number & "]" & Err.Description, "Error")
  Call Library.errorHandle
End Function


'==================================================================================================
Public Function GetHollyday(targetdate As Date, HollydayName As String) As Boolean
    Dim kaerichi As Boolean
    
    kaerichi = False
    HollydayName = ""
    kaerichi = NationalHollydays(targetdate, HollydayName)
    If kaerichi = True Then
        GetHollyday = True
    Else
        HollydayName = ""
        kaerichi = FurikaeKyujitsu(targetdate, HollydayName)
        If kaerichi = True Then
            GetHollyday = True
        Else
            HollydayName = ""
            kaerichi = KokuminnoKyujitsu(targetdate, HollydayName)
            If kaerichi = True Then
                GetHollyday = True
            Else
                HollydayName = ""
                kaerichi = TokubetsunaKyujitsu(targetdate, HollydayName)
                If kaerichi = True Then
                    GetHollyday = True
                Else
                    GetHollyday = False
                End If
            End If
        End If
    End If
End Function
'******************************************************************************
' 祝日判定関数
'******************************************************************************
Public Function NationalHollydays(targetdate As Date, HollydayName As String) As Boolean
  Dim targetyear As Integer
  Dim targetmonth As Integer
  Dim targetDay As Integer
  Dim hantei As Boolean
  
    targetyear = CInt(Format(targetdate, "yyyy"))
    targetmonth = CInt(Format(targetdate, "m"))
    targetDay = CInt(Format(targetdate, "d"))
    hantei = False
    Select Case targetmonth
    Case 1
        If targetyear > 1948 And targetDay = 1 Then
                hantei = True
                HollydayName = "元旦"
        End If
        If targetyear > 1948 Then
            If targetyear < 2000 Then
                If targetDay = 15 Then
                    hantei = True
                    HollydayName = "成人の日"
                End If
            ElseIf CInt(Format(DaiXYoubi(targetyear, 1, 2, 1), "d")) = targetDay Then
                    hantei = True
                    HollydayName = "成人の日"
            End If
        End If
    Case 2
        If targetyear > 1966 Then
                If targetDay = 11 Then
                    hantei = True
                    HollydayName = "建国記念の日"
                End If
         End If
        If targetyear > 2019 Then
            If targetDay = 23 Then
                hantei = True
                HollydayName = "天皇誕生日"
            End If
        End If
    Case 3
        If targetyear > 1948 Then
            If targetDay = Syunbun(targetyear) Then
                hantei = True
                HollydayName = "春分の日"
            End If
        End If

        
    Case 4
        If targetDay = 29 Then
            If targetyear > 1948 Then
                If 1989 > targetyear Then
                    hantei = True
                    HollydayName = "天皇誕生日"
            ElseIf 2007 > targetyear And targetyear > 1988 Then
                    hantei = True
                    HollydayName = "みどりの日"
                Else
                    hantei = True
                    HollydayName = "昭和の日"
                End If
            End If
        End If
    Case 5
        If targetyear > 1948 Then
            If targetDay = 3 Then
                    hantei = True
                    HollydayName = "憲法記念日"
            End If
            If targetDay = 5 Then
                    hantei = True
                    HollydayName = "こどもの日"
            End If
            If targetDay = 4 Then
                If targetyear > 2006 Then
                    hantei = True
                    HollydayName = "みどりの日"
                End If
            End If
        End If
    Case 7
            '海の日[7月の第3月曜日]---------------------------------------------
            '東京オリンピック競技大会・東京パラリンピック競技大会特別措置法等の一部を改正する法律
            If targetyear = 2020 Then
              If targetDay = 23 Then
                hantei = True
                HollydayName = "海の日"
              End If
            ElseIf targetyear = 2021 Then
              If targetDay = 22 Then
                hantei = True
                HollydayName = "海の日"
              End If
              
            '
            ElseIf targetyear > 1995 Then
              If 2004 > targetyear Then
                If targetDay = 20 Then
                    hantei = True
                    HollydayName = "海の日"
                End If
              Else
                If CInt(Format(DaiXYoubi(targetyear, 7, 3, 0), "d")) = targetDay Then
                  hantei = True
                  HollydayName = "海の日"
                End If
              End If
            End If
            
            
            '東京オリンピック競技大会・東京パラリンピック競技大会特別措置法等の一部を改正する法律
            If targetyear = 2020 Then
              If targetDay = 24 Then
                hantei = True
                HollydayName = "スポーツの日"
              End If
              
            ElseIf targetyear = 2021 Then
              If targetDay = 23 Then
                hantei = True
                HollydayName = "スポーツの日"
              End If
            End If
            
    Case 8
            '山の日[8/11]---------------------------------------------
            '東京オリンピック競技大会・東京パラリンピック競技大会特別措置法等の一部を改正する法律
            If targetyear = 2021 Then
              If targetDay = 8 Then
                hantei = True
                HollydayName = "山の日"
              End If
            ElseIf targetyear = 2020 Then
              If targetDay = 10 Then
                hantei = True
                HollydayName = "山の日"
              End If
              
            ElseIf targetyear >= 2016 And targetDay = 11 Then
              hantei = True
              HollydayName = "山の日"
            End If

                
                
                
    Case 9
        If targetyear > 1965 Then
            If 2004 > targetyear Then
                If targetDay = 15 Then
                    hantei = True
                    HollydayName = "敬老の日"
                End If
            Else
                If targetyear > 2003 And CInt(Format(DaiXYoubi(targetyear, 9, 3, 1), "d")) = targetDay Then
                    hantei = True
                    HollydayName = "敬老の日"
                End If
            End If
        End If
        If targetyear > 1947 Then
            If targetDay = Syuubun(targetyear) Then
                hantei = True
                HollydayName = "秋分の日"
            End If
        End If
    Case 10
      '体育の日[10/10→10月の第2月曜日]---------------------------------------------
        If targetyear > 1965 Then
          If 2000 > targetyear Then
            If targetDay = 10 Then
              hantei = True
              HollydayName = "体育の日"
            End If
          ElseIf targetyear > 1999 And targetyear < 2020 Then
            If CInt(Format(DaiXYoubi(targetyear, 10, 2, 1), "d")) = targetDay Then
              hantei = True
              HollydayName = "体育の日"
            End If
          ElseIf targetyear = 2020 Or targetyear = 2021 Then
            '東京オリンピック競技大会・東京パラリンピック競技大会特別措置法等の一部を改正する法律
            '7月に移動
            
          ElseIf targetyear > 2020 Then
            If CInt(Format(DaiXYoubi(targetyear, 10, 2, 1), "d")) = targetDay Then
              hantei = True
              HollydayName = "スポーツの日"
            End If
          End If
        End If
    Case 11
        If targetyear > 1947 Then
            If targetDay = 3 Then
                hantei = True
                HollydayName = "文化の日"
            ElseIf targetDay = 23 Then
                hantei = True
                HollydayName = "勤労感謝の日"
            End If
        End If
    Case 12
        If targetyear > 1988 And targetyear <= 2018 Then
            If targetDay = 23 Then
                hantei = True
                HollydayName = "天皇誕生日"
            End If
        End If
    End Select
    If hantei = True Then
        NationalHollydays = True
    Else
        NationalHollydays = False
    End If
End Function
'******************************************************************************
' 春分の日を求める
'******************************************************************************
Public Function Syunbun(Nen As Integer) As Integer
    
    Syunbun = 0
    If (1899 >= Nen And Nen >= 1851) Then
        Syunbun = Int(19.8277 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If (1979 >= Nen And Nen >= 1900) Then
        Syunbun = Int(20.8357 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If 2099 >= Nen And Nen >= 1980 Then
        Syunbun = Int(20.8431 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If
    If (2150 >= Nen And Nen >= 2100) Then
        Syunbun = Int(21.851 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If

End Function
'******************************************************************************
' 秋分の日を求める
'******************************************************************************
Public Function Syuubun(Nen As Integer) As Integer

    
  Syuubun = 0
  If (1899 >= Nen And Nen >= 1851) Then
      Syuubun = Int(22.2588 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
  End If
  If (1979 >= Nen And Nen >= 1900) Then
      Syuubun = Int(23.2588 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
  End If
  If (2099 >= Nen And Nen >= 1980) Then
      Syuubun = Int(23.2488 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
  End If
  If (2150 >= Nen And Nen >= 2100) Then
      Syuubun = Int(24.2488 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
  End If
End Function
'******************************************************************************
' ある月の第○△曜日が□日であるかを調べる関数。
'******************************************************************************
Public Function DaiXYoubi(y, m, n, Yobi As Integer) As String
    DaiXYoubi = ((9 - Weekday(DateSerial(y, m, 0))) + (n - 1) * 7 + 1)
End Function
'******************************************************************************
' 振替休日かを調べる関数。
'******************************************************************************
Public Function FurikaeKyujitsu(targetdate As Date, HollydayName As String) As Boolean
  Dim lastsunday  As Date
  Dim days As Integer
  Dim hantei As Boolean
  Dim targetyear As Integer, i As Integer
  
  
  HollydayName = ""
  hantei = False
  lastsunday = DateAdd("d", 1 - (Weekday(targetdate)), targetdate)
  days = (Weekday(targetdate) - 1)
  If targetdate > "1973/04/11" Then
      If NationalHollydays(targetdate, HollydayName) = False Then
          If targetyear < 2007 Then
              If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And Weekday(targetdate) = 2 Then
                  HollydayName = "振替休日"
                  FurikaeKyujitsu = True
              Else
                  HollydayName = ""
                  FurikaeKyujitsu = False
              End If
          Else
              If NationalHollydays(lastsunday, HollydayName) = True Then
                  For i = 0 To (days - 1)
                      If NationalHollydays(DateAdd("d", i, lastsunday), HollydayName) = False Then
                          FurikaeKyujitsu = False
                          HollydayName = ""
                          Exit Function
                      End If
                  Next i
                  HollydayName = "振替休日"
                  FurikaeKyujitsu = True
              Else
                  FurikaeKyujitsu = False
                  HollydayName = ""
              End If
          End If
      End If
  End If
End Function
'******************************************************************************
' 国民の休日かを調べる関数。
'******************************************************************************
Public Function KokuminnoKyujitsu(targetdate As Date, HollydayName As String) As Boolean
  Dim targetyear As Integer, i As Integer
    
    HollydayName = ""
    If targetdate > "1985/12/26" Then
        If NationalHollydays(targetdate, HollydayName) = False Then
            If targetyear < 2007 Then
                If FurikaeKyujitsu(targetdate, HollydayName) = False And Weekday(targetdate) <> 1 Then
                    If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And NationalHollydays(DateAdd("d", 1, targetdate), HollydayName) = True Then
                        HollydayName = "国民の休日"
                        KokuminnoKyujitsu = True
                    Else
                        HollydayName = ""
                        KokuminnoKyujitsu = False
                    End If
                Else
                    HollydayName = ""
                    KokuminnoKyujitsu = False
                End If
            Else
                If NationalHollydays(targetdate, HollydayName) = False Then
                    If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And NationalHollydays(DateAdd("d", 1, targetdate), HollydayName) = True Then
                        HollydayName = "国民の休日"
                        KokuminnoKyujitsu = True
                    Else
                        HollydayName = ""
                        KokuminnoKyujitsu = False
                    End If
                Else
                    HollydayName = ""
                    KokuminnoKyujitsu = False
                End If
            End If
        End If
    End If
End Function
'******************************************************************************
' 特別な休日
'******************************************************************************
Public Function TokubetsunaKyujitsu(targetdate As Date, HollydayName As String) As Boolean
    Dim line As Long, endLine As Long
    
    TokubetsunaKyujitsu = False
    If targetdate = "1959/04/10" Then
        HollydayName = "明仁親王結婚の儀"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1989/02/24" Then
        HollydayName = "昭和天皇大喪の礼"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1990/11/12" Then
        HollydayName = "即位礼正殿の儀"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1993/06/09" Then
        HollydayName = "徳仁親王結婚の儀"
        TokubetsunaKyujitsu = True
    End If
    
'    If targetdate = "2020/07/23" Or targetdate = "2021/07/22" Then
'        HollydayName = "海の日"
'        TokubetsunaKyujitsu = True
'    End If
'    If targetdate = "2020/07/24" Or targetdate = "2021/07/23" Then
'        HollydayName = "スポーツの日"
'        TokubetsunaKyujitsu = True
'    End If
'    If targetdate = "2020/08/10" Or targetdate = "2021/08/08" Then
'        HollydayName = "山の日"
'        TokubetsunaKyujitsu = True
'    End If
    
    
'    '会社指定休日の設定
'    endLine = sheetsetting.Cells(Rows.count, Library.getColumnNo(LadexsetVal("cell_CompanyHoliday"))).End(xlUp).Row
'    For line = 3 To endLine
'      If targetdate = sheetsetting.Range(LadexsetVal("cell_CompanyHoliday") & line) Then
'          HollydayName = "会社指定休日"
'          TokubetsunaKyujitsu = True
'      End If
'    Next
    
    
    
    
    
End Function
