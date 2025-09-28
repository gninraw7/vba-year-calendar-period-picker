Attribute VB_Name = "modCallBackRibbon"
' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
'
'' 제작 : 임길재 (M:O1O-2716-0735, E-Mail: gninraw7@naver.com, kiljae.lim@gmail.com)
'' 2025-09-08 최초 작성
'' 2025-09-15 Addin 작성

Option Explicit

' Ribbon 핸들 & 상태
Public gRibbon As IRibbonUI

Public Const Const_PeriodPicker_Menu As String = "PeriodPicker_Menu"

' 이미 선언되어 있지 않다면 모듈 상단 등 공용 범위에 선언
Public gHolidaySet As Object   ' 'Scripting.Dictionary'

Private Const REG_APP As String = "PeriodPicker"
Private Const REG_SEC As String = "Holidays"

'----------------------
' 초기화 / Ribbon 기본
'----------------------
Public Sub PP_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

'Callback for btnViewYearCalendar onAction
Sub onViewYearCalendar(control As IRibbonControl)
    Show_YearCalendar
End Sub

Sub Show_YearCalendar()
    Dim f As New frmYearCalendar
    f.SetTargetRange Selection         ' 선택영역이 없으면 내부에서 Selection 읽음
    f.Show vbModeless
End Sub

'Callback for btnSetHolidays onAction
Sub onSetHolidays(control As IRibbonControl)
    frmSetHolidays.Show vbModeless
End Sub

'===========================
Sub Delete_PeriodPicker_Menu()
    On Error Resume Next
    Remove_CommandBar
    Application.CommandBars(Const_PeriodPicker_Menu).Delete
    On Error GoTo 0
End Sub

Sub Gen_PeriodPicker_CommandBar()
    Dim L_CommandBar As CommandBar
    
    Call Remove_CommandBar
    Call Reset_CommandBars
    
    Set L_CommandBar = Application.CommandBars("Cell")
    
    With L_CommandBar.Controls.add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "Show_YearCalendar"
        .Caption = "Year Calendar"
        .Tag = "PeriodPicker_Cell_Control_Tag"
    End With

    With L_CommandBar.Controls.add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "InsertYesterday"
        .Caption = "어제"
        .Tag = "PeriodPicker_Cell_Control_Tag"
    End With
    
    With L_CommandBar.Controls.add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "InsertToday"
        .Caption = "오늘"
        .Tag = "PeriodPicker_Cell_Control_Tag"
    End With
    
    L_CommandBar.Controls(4).BeginGroup = True
End Sub

Sub Reset_CommandBars()
    On Error Resume Next
    CommandBars("Cell").Reset
    On Error GoTo 0
End Sub

Sub Remove_CommandBar()
    Dim L_CommandBar As CommandBar
    Dim L_CommandBarControl As CommandBarControl
    On Error Resume Next
    Set L_CommandBar = Application.CommandBars("Cell")
    For Each L_CommandBarControl In L_CommandBar.Controls
        If L_CommandBarControl.Tag = "PeriodPicker_Cell_Control_Tag" Then
            L_CommandBarControl.Delete
        End If
    Next L_CommandBarControl
    
    On Error GoTo 0
End Sub

Sub InsertToday()
    ' 선택된 셀에 오늘 날짜 입력
    If Not Selection Is Nothing Then
        Selection.Value = Date
    End If
End Sub

Sub InsertYesterday()
   ' 선택된 셀에 어제 날짜 입력
    If Not Selection Is Nothing Then
        Selection.Value = Date - 1
    End If
End Sub


'선택 영역이 1셀이면 "YYYY-MM-DD ~ YYYY-MM-DD" 문자열로 채움
'2셀 이상이면 첫 셀=From, 그 오른쪽 셀=To(날짜 형식)
'
'(2) 다른 폼의 TextBox 두 개로 반환
' 예: frmOrder 에 TextBox txtFrom, txtTo 가 있다고 가정
'Sub ShowYearCalendar_ToOtherFormTextBoxes()
'    Dim f As New frmYearCalendar
'    f.SetTargetTextBoxes frmOrder.txtFrom, frmOrder.txtTo
'    f.Show
'End Sub


' ================== 메인: 모든 연도 → gHolidaySet 로드 ==================
Public Sub LoadAllHolidaysIntoGlobalSet()
    ' 1) dict 준비
    If gHolidaySet Is Nothing Then
        Set gHolidaySet = CreateObject("Scripting.Dictionary")
    Else
        gHolidaySet.RemoveAll
    End If
    ' 날짜 키는 숫자 비교이므로 CompareMode 영향 없음

    Dim kv As Variant, i As Long, found As Boolean
    On Error Resume Next
    kv = GetAllSettings(REG_APP, REG_SEC)   ' 2차원 배열([row, 0]=key, [row, 1]=value)
    On Error GoTo 0

    If IsArray(kv) Then
        ' 2) 섹션에 존재하는 모든 (연도키, 원문) 처리
        For i = LBound(kv, 1) To UBound(kv, 1)
            AddHolidayRawToDict CStr(kv(i, 1)), gHolidaySet  ' value=원문
        Next
        found = True
    End If

    ' 3) 폴백: 섹션이 없거나 읽기 실패 시 연도 범위 스캔
    If Not found Then
        Dim y As Long, raw As String
        For y = 1900 To 2099
            raw = GetSetting(REG_APP, REG_SEC, CStr(y), "")
            If Len(raw) > 0 Then AddHolidayRawToDict raw, gHolidaySet
        Next
    End If

    ' 원하면 디버그 카운트 출력
    'Debug.Print "gHolidaySet.Count=" & gHolidaySet.Count
End Sub

' ================== 원문 파싱: "yyyy-mm-dd|휴일명" 줄들 → dict(Date → 이름) ==================
Private Sub AddHolidayRawToDict(ByVal raw As String, ByRef dict As Object)
    If Len(raw) = 0 Then Exit Sub

    Dim lines() As String, i As Long, s As String
    lines = Split(raw, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Len(s) = 0 Then GoTo ContinueNext

        Dim a() As String, ymd As String, nm As String, dT As Date, ok As Boolean
        a = Split(s, "|")
        ymd = Trim$(a(0))
        If UBound(a) >= 1 Then nm = CStr(a(1)) Else nm = ""

        ok = TryParseYMDToDate(ymd, dT)
        If ok Then
            ' 같은 날짜가 이미 있으면 마지막 값을 우선으로 덮어씀
            dict(CDate(Int(CDbl(dT)))) = nm   ' 자정으로 정규화(날짜 키 안정화)
        End If
ContinueNext:
    Next
End Sub

Sub TEST_LoadAll()
    Call LoadAllHolidaysIntoGlobalSet
    Dim k As Variant
    For Each k In gHolidaySet.keys
        Debug.Print Format$(CDate(k), "yyyy-mm-dd"), gHolidaySet(k)
    Next
End Sub

Public Sub SeedHolidays_InitialInstall(Optional ByVal forceOverwrite As Boolean = False)
    Dim y As Long, cur As String, def As String, wrote As Long

    For y = 2004 To 2027
        cur = GetSetting(REG_APP, REG_SEC, CStr(y), "")
        If forceOverwrite Or Len(cur) = 0 Then
            def = DefaultHolidayLines(y)   ' ▼ 아래 함수에 내장된 기본값 사용
            If Len(def) > 0 Then
                SaveSetting REG_APP, REG_SEC, CStr(y), def
                wrote = wrote + 1
            End If
        End If
    Next

'    MsgBox IIf(forceOverwrite, "[강제]", "[비어있는 연도만]") & _
'           " 공휴일 시드 완료: " & wrote & "개 연도", vbInformation
End Sub

'=========================
' ③ 내장 기본값(상수화된 원본)  ← 여기 바디를 “생성기 출력물”로 교체하세요
'     포맷: 레지스트리에 저장해 두던 raw 문자열 그대로
'     (줄당 "YYYY-MM-DD<구분자>휴일명" 형태, 구분자는 기존 저장 포맷 유지)
'=========================
Public Function DefaultHolidayLines(ByVal y As Long) As String
    Select Case y
        Case 2004: DefaultHolidayLines = "2004-01-01|신정" & vbCrLf & "2004-01-21|설날" & vbCrLf & "2004-01-22|설날" & vbCrLf & "2004-01-23|설날" & vbCrLf & "2004-03-01|삼일절" & vbCrLf & "2004-04-05|식목일" & vbCrLf & "2004-05-05|어린이날" & vbCrLf & "2004-05-26|석가탄신일" & vbCrLf & "2004-06-06|현충일" & vbCrLf & "2004-07-17|제헌절" & vbCrLf & "2004-08-15|광복절" & vbCrLf & "2004-09-27|추석" & vbCrLf & "2004-09-28|추석" & vbCrLf & "2004-09-29|추석" & vbCrLf & "2004-10-03|개천절" & vbCrLf & "2004-12-25|기독탄신일" & vbCrLf & ""
        Case 2005: DefaultHolidayLines = "2005-01-01|신정" & vbCrLf & "2005-02-08|설날" & vbCrLf & "2005-02-09|설날" & vbCrLf & "2005-02-10|설날" & vbCrLf & "2005-03-01|삼일절" & vbCrLf & "2005-04-05|식목일" & vbCrLf & "2005-05-05|어린이날" & vbCrLf & "2005-05-15|석가탄신일" & vbCrLf & "2005-06-06|현충일" & vbCrLf & "2005-07-17|제헌절" & vbCrLf & "2005-08-15|광복절" & vbCrLf & "2005-09-17|추석" & vbCrLf & "2005-09-18|추석" & vbCrLf & "2005-09-19|추석" & vbCrLf & "2005-10-03|개천절" & vbCrLf & "2005-12-25|기독탄신일" & vbCrLf & ""
        Case 2006: DefaultHolidayLines = "2006-01-01|신정" & vbCrLf & "2006-01-28|설날" & vbCrLf & "2006-01-29|설날" & vbCrLf & "2006-01-30|설날" & vbCrLf & "2006-03-01|삼일절" & vbCrLf & "2006-05-05|석가탄신일" & vbCrLf & "2006-06-06|현충일" & vbCrLf & "2006-07-17|제헌절" & vbCrLf & "2006-08-15|광복절" & vbCrLf & "2006-10-03|개천절" & vbCrLf & "2006-10-05|추석" & vbCrLf & "2006-10-06|추석" & vbCrLf & "2006-10-07|추석" & vbCrLf & "2006-12-25|기독탄신일" & vbCrLf & ""
        Case 2007: DefaultHolidayLines = "2007-01-01|신정" & vbCrLf & "2007-02-17|설날" & vbCrLf & "2007-02-18|설날" & vbCrLf & "2007-02-19|설날" & vbCrLf & "2007-03-01|삼일절" & vbCrLf & "2007-05-05|어린이날" & vbCrLf & "2007-05-24|석가탄신일" & vbCrLf & "2007-06-06|현충일" & vbCrLf & "2007-07-17|제헌절" & vbCrLf & "2007-08-15|광복절" & vbCrLf & "2007-09-24|추석" & vbCrLf & "2007-09-25|추석" & vbCrLf & "2007-09-26|추석" & vbCrLf & "2007-10-03|개천절" & vbCrLf & "2007-12-25|기독탄신일" & vbCrLf & ""
        Case 2008: DefaultHolidayLines = "2008-01-01|신정" & vbCrLf & "2008-02-06|설날" & vbCrLf & "2008-02-07|설날" & vbCrLf & "2008-02-08|설날" & vbCrLf & "2008-03-01|삼일절" & vbCrLf & "2008-05-05|어린이날" & vbCrLf & "2008-05-12|석가탄신일" & vbCrLf & "2008-06-06|현충일" & vbCrLf & "2008-08-15|광복절" & vbCrLf & "2008-09-13|추석" & vbCrLf & "2008-09-14|추석" & vbCrLf & "2008-09-15|추석" & vbCrLf & "2008-10-03|개천절" & vbCrLf & "2008-12-25|기독탄신일" & vbCrLf & ""
        Case 2009: DefaultHolidayLines = "2009-01-01|신정" & vbCrLf & "2009-01-25|설날" & vbCrLf & "2009-01-26|설날" & vbCrLf & "2009-01-27|설날" & vbCrLf & "2009-03-01|삼일절" & vbCrLf & "2009-05-02|석가탄신일" & vbCrLf & "2009-05-05|어린이날" & vbCrLf & "2009-06-06|현충일" & vbCrLf & "2009-08-15|광복절" & vbCrLf & "2009-10-02|추석" & vbCrLf & "2009-10-03|개천절" & vbCrLf & "2009-10-04|추석" & vbCrLf & "2009-12-25|기독탄신일" & vbCrLf & ""
        Case 2010: DefaultHolidayLines = "2010-01-01|신정" & vbCrLf & "2010-02-13|설날" & vbCrLf & "2010-02-14|설날" & vbCrLf & "2010-02-15|설날" & vbCrLf & "2010-03-01|삼일절" & vbCrLf & "2010-05-05|어린이날" & vbCrLf & "2010-05-21|석가탄신일" & vbCrLf & "2010-06-06|현충일" & vbCrLf & "2010-08-15|광복절" & vbCrLf & "2010-09-21|추석" & vbCrLf & "2010-09-22|추석" & vbCrLf & "2010-09-23|추석" & vbCrLf & "2010-10-03|개천절" & vbCrLf & "2010-12-25|기독탄신일" & vbCrLf & ""
        Case 2011: DefaultHolidayLines = "2011-01-01|신정" & vbCrLf & "2011-02-02|설날" & vbCrLf & "2011-02-03|설날" & vbCrLf & "2011-02-04|설날" & vbCrLf & "2011-03-01|삼일절" & vbCrLf & "2011-05-05|어린이날" & vbCrLf & "2011-05-10|석가탄신일" & vbCrLf & "2011-06-06|현충일" & vbCrLf & "2011-08-15|광복절" & vbCrLf & "2011-09-11|추석" & vbCrLf & "2011-09-12|추석" & vbCrLf & "2011-09-13|추석" & vbCrLf & "2011-10-03|개천절" & vbCrLf & "2011-12-25|기독탄신일" & vbCrLf & ""
        Case 2012: DefaultHolidayLines = "2012-01-01|신정" & vbCrLf & "2012-01-22|설날" & vbCrLf & "2012-01-23|설날" & vbCrLf & "2012-01-24|설날" & vbCrLf & "2012-03-01|삼일절" & vbCrLf & "2012-04-11|국회의원선거일" & vbCrLf & "2012-05-05|어린이날" & vbCrLf & "2012-05-28|석가탄신일" & vbCrLf & "2012-06-06|현충일" & vbCrLf & "2012-08-15|광복절" & vbCrLf & "2012-09-29|추석" & vbCrLf & "2012-09-30|추석" & vbCrLf & "2012-10-01|추석" & vbCrLf & "2012-10-03|개천절" & vbCrLf & "2012-12-19|대통령선거일" & vbCrLf & "2012-12-25|기독탄신일" & vbCrLf & ""
        Case 2013: DefaultHolidayLines = "2013-01-01|신정" & vbCrLf & "2013-02-09|설날" & vbCrLf & "2013-02-10|설날" & vbCrLf & "2013-02-11|설날" & vbCrLf & "2013-03-01|삼일절" & vbCrLf & "2013-05-05|어린이 날" & vbCrLf & "2013-05-17|석가탄신일" & vbCrLf & "2013-06-06|현충일" & vbCrLf & "2013-08-15|광복절" & vbCrLf & "2013-09-18|추석" & vbCrLf & "2013-09-19|추석" & vbCrLf & "2013-09-20|추석" & vbCrLf & "2013-10-03|개천절" & vbCrLf & "2013-10-09|한글날" & vbCrLf & "2013-12-25|기독탄신일" & vbCrLf & ""
        Case 2014: DefaultHolidayLines = "2014-01-01|신정" & vbCrLf & "2014-01-30|설날" & vbCrLf & "2014-01-31|설날" & vbCrLf & "2014-02-01|설날" & vbCrLf & "2014-03-01|삼일절" & vbCrLf & "2014-05-05|어린이날" & vbCrLf & "2014-05-06|석가탄신일" & vbCrLf & "2014-06-04|동시지방선거일" & vbCrLf & "2014-06-06|현충일" & vbCrLf & "2014-08-15|광복절" & vbCrLf & "2014-09-07|추석" & vbCrLf & "2014-09-08|추석" & vbCrLf & "2014-09-09|추석" & vbCrLf & "2014-09-10|대체공휴일" & vbCrLf & "2014-10-03|개천절" & vbCrLf & "2014-10-09|한글날" & vbCrLf & "2014-12-25|기독탄신일" & vbCrLf & ""
        Case 2015: DefaultHolidayLines = "2015-01-01|신정" & vbCrLf & "2015-02-18|설날" & vbCrLf & "2015-02-19|설날" & vbCrLf & "2015-02-20|설날" & vbCrLf & "2015-03-01|삼일절" & vbCrLf & "2015-05-05|어린이날" & vbCrLf & "2015-05-25|석가탄신일" & vbCrLf & "2015-06-06|현충일" & vbCrLf & "2015-08-15|광복절" & vbCrLf & "2015-09-26|추석" & vbCrLf & "2015-09-27|추석" & vbCrLf & "2015-09-28|추석" & vbCrLf & "2015-09-29|대체공휴일" & vbCrLf & "2015-10-03|개천절" & vbCrLf & "2015-10-09|한글날" & vbCrLf & "2015-12-25|기독탄신일" & vbCrLf & ""
        Case 2016: DefaultHolidayLines = "2016-01-01|신정" & vbCrLf & "2016-02-07|설날" & vbCrLf & "2016-02-08|설날" & vbCrLf & "2016-02-09|설날" & vbCrLf & "2016-02-10|대체공휴일" & vbCrLf & "2016-03-01|삼일절" & vbCrLf & "2016-04-13|국회의원선거일" & vbCrLf & "2016-05-05|어린이날" & vbCrLf & "2016-05-06|임시공휴일" & vbCrLf & "2016-05-14|석가탄신일" & vbCrLf & "2016-06-06|현충일" & vbCrLf & "2016-08-15|광복절" & vbCrLf & "2016-09-14|추석" & vbCrLf & "2016-09-15|추석" & vbCrLf & "2016-09-16|추석" & vbCrLf & "2016-10-03|개천절" & vbCrLf & "2016-10-09|한글날" & vbCrLf & "2016-12-25|기독탄신일" & vbCrLf & ""
        Case 2017: DefaultHolidayLines = "2017-01-01|신정" & vbCrLf & "2017-01-27|설날" & vbCrLf & "2017-01-28|설날" & vbCrLf & "2017-01-29|설날" & vbCrLf & "2017-01-30|대체공휴일" & vbCrLf & "2017-03-01|삼일절" & vbCrLf & "2017-05-03|석가탄신일" & vbCrLf & "2017-05-05|어린이날" & vbCrLf & "2017-05-09|대통령선거일" & vbCrLf & "2017-06-06|현충일" & vbCrLf & "2017-08-15|광복절" & vbCrLf & "2017-10-02|임시공휴일" & vbCrLf & "2017-10-03|추석" & vbCrLf & "2017-10-04|추석" & vbCrLf & "2017-10-05|추석" & vbCrLf & "2017-10-06|대체공휴일" & vbCrLf & "2017-10-09|한글날" & vbCrLf & "2017-12-25|기독탄신일" & vbCrLf & ""
        Case 2018: DefaultHolidayLines = "2018-01-01|1월1일" & vbCrLf & "2018-02-15|설날" & vbCrLf & "2018-02-16|설날" & vbCrLf & "2018-02-17|설날" & vbCrLf & "2018-03-01|삼일절" & vbCrLf & "2018-05-05|어린이날" & vbCrLf & "2018-05-07|대체휴무일" & vbCrLf & "2018-05-22|부처님오신날" & vbCrLf & "2018-06-06|현충일" & vbCrLf & "2018-06-13|전국동시지방선거" & vbCrLf & "2018-08-15|광복절" & vbCrLf & "2018-09-23|추석" & vbCrLf & "2018-09-24|추석" & vbCrLf & "2018-09-25|추석" & vbCrLf & "2018-09-26|대체휴무일" & vbCrLf & "2018-10-03|개천절" & vbCrLf & "2018-10-09|한글날" & vbCrLf & "2018-12-25|기독탄신일" & vbCrLf & ""
        Case 2019: DefaultHolidayLines = "2019-01-01|1월1일" & vbCrLf & "2019-02-04|설날" & vbCrLf & "2019-02-05|설날" & vbCrLf & "2019-02-06|설날" & vbCrLf & "2019-03-01|삼일절" & vbCrLf & "2019-05-05|어린이날" & vbCrLf & "2019-05-06|대체공휴일" & vbCrLf & "2019-05-12|부처님오신날" & vbCrLf & "2019-06-06|현충일" & vbCrLf & "2019-08-15|광복절" & vbCrLf & "2019-09-12|추석" & vbCrLf & "2019-09-13|추석" & vbCrLf & "2019-09-14|추석" & vbCrLf & "2019-10-03|개천절" & vbCrLf & "2019-10-09|한글날" & vbCrLf & "2019-12-25|기독탄신일" & vbCrLf & ""
        Case 2020: DefaultHolidayLines = "2020-01-01|1월1일" & vbCrLf & "2020-01-24|설날" & vbCrLf & "2020-01-25|설날" & vbCrLf & "2020-01-26|설날" & vbCrLf & "2020-01-27|대체공휴일" & vbCrLf & "2020-03-01|삼일절" & vbCrLf & "2020-04-15|제21대 국회의원선거" & vbCrLf & "2020-04-30|부처님오신날" & vbCrLf & "2020-05-05|어린이날" & vbCrLf & "2020-06-06|현충일" & vbCrLf & "2020-08-15|광복절" & vbCrLf & "2020-08-17|임시공휴일" & vbCrLf & "2020-09-30|추석" & vbCrLf & "2020-10-01|추석" & vbCrLf & "2020-10-02|추석" & vbCrLf & "2020-10-03|개천절" & vbCrLf & "2020-10-09|한글날" & vbCrLf & "2020-12-25|기독탄신일" & vbCrLf & ""
        Case 2021: DefaultHolidayLines = "2021-01-01|1월1일" & vbCrLf & "2021-02-11|설날" & vbCrLf & "2021-02-12|설날" & vbCrLf & "2021-02-13|설날" & vbCrLf & "2021-03-01|삼일절" & vbCrLf & "2021-05-05|어린이날" & vbCrLf & "2021-05-19|부처님오신날" & vbCrLf & "2021-06-06|현충일" & vbCrLf & "2021-08-15|광복절" & vbCrLf & "2021-08-16|대체공휴일" & vbCrLf & "2021-09-20|추석" & vbCrLf & "2021-09-21|추석" & vbCrLf & "2021-09-22|추석" & vbCrLf & "2021-10-03|개천절" & vbCrLf & "2021-10-04|대체공휴일" & vbCrLf & "2021-10-09|한글날" & vbCrLf & "2021-10-11|대체공휴일" & vbCrLf & "2021-12-25|기독탄신일" & vbCrLf & ""
        Case 2022: DefaultHolidayLines = "2022-01-01|1월1일" & vbCrLf & "2022-01-31|설날" & vbCrLf & "2022-02-01|설날" & vbCrLf & "2022-02-02|설날" & vbCrLf & "2022-03-01|삼일절" & vbCrLf & "2022-03-09|대통령선거일" & vbCrLf & "2022-05-05|어린이날" & vbCrLf & "2022-05-08|부처님오신날" & vbCrLf & "2022-06-01|전국동시지방선거" & vbCrLf & "2022-06-06|현충일" & vbCrLf & "2022-08-15|광복절" & vbCrLf & "2022-09-09|추석" & vbCrLf & "2022-09-10|추석" & vbCrLf & "2022-09-11|추석" & vbCrLf & "2022-09-12|대체공휴일" & vbCrLf & "2022-10-03|개천절" & vbCrLf & "2022-10-09|한글날" & vbCrLf & "2022-10-10|대체공휴일" & vbCrLf & "2022-12-25|기독탄신일" & vbCrLf & ""
        Case 2023: DefaultHolidayLines = "2023-01-01|1월1일" & vbCrLf & "2023-01-21|설날" & vbCrLf & "2023-01-22|설날" & vbCrLf & "2023-01-23|설날" & vbCrLf & "2023-01-24|대체공휴일" & vbCrLf & "2023-03-01|삼일절" & vbCrLf & "2023-05-05|어린이날" & vbCrLf & "2023-05-27|부처님오신날" & vbCrLf & "2023-05-29|대체공휴일" & vbCrLf & "2023-06-06|현충일" & vbCrLf & "2023-08-15|광복절" & vbCrLf & "2023-09-28|추석" & vbCrLf & "2023-09-29|추석" & vbCrLf & "2023-09-30|추석" & vbCrLf & "2023-10-02|임시공휴일" & vbCrLf & "2023-10-03|개천절" & vbCrLf & "2023-10-09|한글날" & vbCrLf & "2023-12-25|기독탄신일" & vbCrLf & ""
        Case 2024: DefaultHolidayLines = "2024-01-01|1월1일" & vbCrLf & "2024-02-09|설날" & vbCrLf & "2024-02-10|설날" & vbCrLf & "2024-02-11|설날" & vbCrLf & "2024-02-12|대체공휴일(설날)" & vbCrLf & "2024-03-01|삼일절" & vbCrLf & "2024-04-10|국회의원선거" & vbCrLf & "2024-05-05|어린이날" & vbCrLf & "2024-05-06|대체공휴일(어린이날)" & vbCrLf & "2024-05-15|부처님오신날" & vbCrLf & "2024-06-06|현충일" & vbCrLf & "2024-08-15|광복절" & vbCrLf & "2024-09-16|추석" & vbCrLf & "2024-09-17|추석" & vbCrLf & "2024-09-18|추석" & vbCrLf & "2024-10-01|임시공휴일" & vbCrLf & "2024-10-03|개천절" & vbCrLf & "2024-10-09|한글날" & vbCrLf & "2024-12-25|기독탄신일" & vbCrLf & ""
        Case 2025: DefaultHolidayLines = "2025-01-01|1월1일" & vbCrLf & "2025-01-27|임시공휴일" & vbCrLf & "2025-01-28|설날" & vbCrLf & "2025-01-29|설날" & vbCrLf & "2025-01-30|설날" & vbCrLf & "2025-03-01|삼일절" & vbCrLf & "2025-03-03|대체공휴일" & vbCrLf & "2025-05-05|부처님오신날" & vbCrLf & "2025-05-06|대체공휴일" & vbCrLf & "2025-06-03|임시공휴일(제21대 대통령 선거)" & vbCrLf & "2025-06-06|현충일" & vbCrLf & "2025-08-15|광복절" & vbCrLf & "2025-10-03|개천절" & vbCrLf & "2025-10-05|추석" & vbCrLf & "2025-10-06|추석" & vbCrLf & "2025-10-07|추석" & vbCrLf & "2025-10-08|대체공휴일" & vbCrLf & "2025-10-09|한글날" & vbCrLf & "2025-12-25|기독탄신일" & vbCrLf & ""
        Case 2026: DefaultHolidayLines = "2026-01-01|1월1일" & vbCrLf & "2026-02-16|설날" & vbCrLf & "2026-02-17|설날" & vbCrLf & "2026-02-18|설날" & vbCrLf & "2026-03-01|삼일절" & vbCrLf & "2026-03-02|대체공휴일(삼일절)" & vbCrLf & "2026-05-05|어린이날" & vbCrLf & "2026-05-24|부처님오신날" & vbCrLf & "2026-05-25|대체공휴일(부처님오신날)" & vbCrLf & "2026-06-03|전국동시지방선거" & vbCrLf & "2026-06-06|현충일" & vbCrLf & "2026-08-15|광복절" & vbCrLf & "2026-08-17|대체공휴일(광복절)" & vbCrLf & "2026-09-24|추석" & vbCrLf & "2026-09-25|추석" & vbCrLf & "2026-09-26|추석" & vbCrLf & "2026-10-03|개천절" & vbCrLf & "2026-10-05|대체공휴일(개천절)" & vbCrLf & "2026-10-09|한글날" & vbCrLf & "2026-12-25|기독탄신일" & vbCrLf & ""
        Case 2027: DefaultHolidayLines = "2027-01-01|1월1일" & vbCrLf & "2027-02-06|설날" & vbCrLf & "2027-02-07|설날" & vbCrLf & "2027-02-08|설날" & vbCrLf & "2027-02-09|대체공휴일(설날)" & vbCrLf & "2027-03-01|삼일절" & vbCrLf & "2027-05-05|어린이날" & vbCrLf & "2027-05-13|부처님오신날" & vbCrLf & "2027-06-06|현충일" & vbCrLf & "2027-08-15|광복절" & vbCrLf & "2027-08-16|대체공휴일(광복절)" & vbCrLf & "2027-09-14|추석" & vbCrLf & "2027-09-15|추석" & vbCrLf & "2027-09-16|추석" & vbCrLf & "2027-10-03|개천절" & vbCrLf & "2027-10-04|대체공휴일(개천절)" & vbCrLf & "2027-10-09|한글날" & vbCrLf & "2027-10-11|대체공휴일(한글날)" & vbCrLf & "2027-12-25|기독탄신일" & vbCrLf & "2027-12-27|대체공휴일(기독탄신일)" & vbCrLf & ""
        Case Else: DefaultHolidayLines = ""
    End Select
End Function


