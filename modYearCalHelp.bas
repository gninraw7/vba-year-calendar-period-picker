Attribute VB_Name = "modYearCalHelp"
' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.

Option Explicit

' 내부 사용: 현재 언어(K/E) 조회
Private Function GetCurrentLangFromRegistry() As String
    ' frmYearCalendar가 쓰는 키와 동일하게 접근
    Dim appName As String, secI18N As String
    appName = "PeriodPicker"
    secI18N = "I18N"
    GetCurrentLangFromRegistry = UCase$(GetSetting(appName, secI18N, "Lang", "K"))
End Function

' (호환용) 기존 폼이 호출하던 이름 유지
' → 내부적으로 현재 언어를 감지해서 한/영 매뉴얼을 반환
Public Function BuildYearCalendarManual() As String
    BuildYearCalendarManual = BuildYearCalendarManual_Current()
End Function

' 현재 언어 기준으로 반환(K/E)
Public Function BuildYearCalendarManual_Current() As String
    If GetCurrentLangFromRegistry() = "E" Then
        BuildYearCalendarManual_Current = BuildYearCalendarManual_EN()
    Else
        BuildYearCalendarManual_Current = BuildYearCalendarManual_KO()
    End If
End Function

' ─────────────────────────────────────────────────────────────
' 한국어 매뉴얼
' ─────────────────────────────────────────────────────────────
Public Function BuildYearCalendarManual_KO() As String
    Dim t As String, NL As String: NL = vbCrLf

    t = t & "★ Year Calendar - 사용법" & NL
    t = t & "────────────────────────────────" & NL & NL

    t = t & "■ 목적" & NL
    t = t & "- 1년 달력을 4×3 그리드로 표시하고, 기간(From~To)을 빠르게 선택/반환합니다." & NL
    t = t & "- 언어/날짜서식/요일시작/요일명 스타일은 설정으로 변경 가능하며 레지스트리에 저장됩니다." & NL & NL

    t = t & "■ 기본 조작" & NL
    t = t & "- 일자 셀 좌클릭=From, 우클릭=To, 더블클릭=즉시반환(끝점 보정)" & NL
    t = t & "- 월 타이틀 좌클릭=해당월 1일(From), 우클릭=말일(To), 더블클릭=그 달 전체" & NL
    t = t & "- [적용]으로 반환, [Clear]로 초기화, [닫기]로 종료" & NL & NL

    t = t & "■ 키보드/휠" & NL
    t = t & "- PageUp/Down: ±1년 | Shift: ±10년 | Ctrl: ±5년 | Shift+Ctrl: ±15년" & NL
    t = t & "- End: 오늘의 연도로 이동, Enter: 적용, Esc: 닫기, Del: 초기화" & NL
    t = t & "- 마우스휠: 연도 ±1 (Shift/Ctrl 조합은 위와 동일 배수)" & NL
    t = t & "- 쉼표(,): From=오늘  |  마침표(.): To=오늘" & NL & NL

    t = t & "■ 레이아웃" & NL
    t = t & "- 1→12, 현재월을 시작/끝에, 현재월 @Slot(1~12) 모드 지원" & NL
    t = t & "- 현재월 블록은 연한 파란 배경으로 하이라이트" & NL & NL

    t = t & "■ 설정" & NL
    t = t & "- 언어(K/E), 날짜서식, 월 타이틀 서식(언어별), 요일 시작(일/월), 요일명 스타일(약칭/풀네임)" & NL
    t = t & "- 부적합 월 타이틀 서식(일자 토큰 포함)은 기본값으로 자동 보정" & NL & NL

    t = t & "■ 범위 하이라이트/영업일" & NL
    t = t & "- 범위 칠하기는 연노랑, [범위표시에 휴일 제외] 선택 시 토/일/공휴일 배제" & NL
    t = t & "- 하단 요약: (영업일 / 총일수), 공휴일은 gHolidaySet 기반" & NL & NL

    t = t & "■ Export" & NL
    t = t & "- 새 시트에 레이아웃/연도/선택범위 요약과 함께 달력을 출력" & NL
    t = t & "- 휴일은 셀 코멘트로 이름 표시, 오늘/주말/범위 색 적용" & NL & NL

'    t = t & "■ 레지스트리(저장 위치)" & NL
'    t = t & "- App: MonthlyCalendar" & NL
'    t = t & "  · I18N\\Lang, I18N\\WeekStart, I18N\\WeekNameStyle" & NL
'    t = t & "  · Format\\Date, Format\\MonthTitle_K/E" & NL
'    t = t & "  · Layout\\LayoutMode/AnchorSlot/Apply_From_Only/Exclude_Non_Biz/Show_Keep" & NL & NL

    t = t & "■ 팁/알림" & NL
    t = t & "- 쉼표/마침표 키는 VK 코드(188/190) 사용 권장(환경별 차이 방지)" & NL
    t = t & "- 월요일 시작에서도 일요일만 빨강, 토요일만 파랑으로 표시됩니다." & NL

    BuildYearCalendarManual_KO = t
End Function

' ─────────────────────────────────────────────────────────────
' English manual
' ─────────────────────────────────────────────────────────────
Public Function BuildYearCalendarManual_EN() As String
    Dim t As String, NL As String: NL = vbCrLf

    t = t & "★ Year Calendar - User Guide" & NL
    t = t & "────────────────────────────────" & NL & NL

    t = t & "■ Purpose" & NL
    t = t & "- Display a full-year calendar in a 4×3 grid and quickly select a date range (From?To)." & NL
    t = t & "- Language/date format/week start/day name style are configurable and persisted in the registry." & NL & NL

    t = t & "■ Basic Operations" & NL
    t = t & "- Day cell Left-Click = From, Right-Click = To, Double-Click = return immediately (auto-order)." & NL
    t = t & "- Month title Left-Click = 1st day (From), Right-Click = last day (To), Double-Click = whole month." & NL
    t = t & "- Use [Apply] to return the range, [Clear] to reset, [Close] to exit." & NL & NL

    t = t & "■ Keyboard & Wheel" & NL
    t = t & "- PageUp/Down: ±1 year | Shift: ±10 | Ctrl: ±5 | Shift+Ctrl: ±15" & NL
    t = t & "- End: jump to current year, Enter: Apply, Esc: Close, Del: Reset" & NL
    t = t & "- Mouse wheel: ±1 year (same multipliers with Shift/Ctrl)." & NL
    t = t & "- Comma ( , ): From = Today  |  Period ( . ): To = Today" & NL & NL

    t = t & "■ Layout" & NL
    t = t & "- Modes: 1→12, Current First, Current Last, Current @Slot (1?12)." & NL
    t = t & "- The current month block is gently highlighted in the background." & NL & NL

    t = t & "■ Settings" & NL
    t = t & "- Language (K/E), Date format, Month title format (per language)," & NL
    t = t & "  Week start (Sun/Mon), Day name style (Short/Full)." & NL
    t = t & "- Invalid month-title formats (containing day tokens) are auto-sanitized to defaults." & NL & NL

    t = t & "■ Range Highlight & Business Days" & NL
    t = t & "- Selected range is painted in light yellow; option to exclude Sat/Sun/Holidays from painting." & NL
    t = t & "- Summary shows (Business days / Total days). Holidays rely on gHolidaySet." & NL & NL

    t = t & "■ Export" & NL
    t = t & "- Prints the current layout/year/range summary to a new worksheet." & NL
    t = t & "- Holiday names are inserted as cell comments; weekend/today/range coloring retained." & NL & NL

'    t = t & "■ Registry (Persistence)" & NL
'    t = t & "- App: MonthlyCalendar" & NL
'    t = t & "  · I18N\\Lang, I18N\\WeekStart, I18N\\WeekNameStyle" & NL
'    t = t & "  · Format\\Date, Format\\MonthTitle_K/E" & NL
'    t = t & "  · Layout\\LayoutMode, AnchorSlot, Apply_From_Only, Exclude_Non_Biz, Show_Keep" & NL & NL

    t = t & "■ Tips / Notes" & NL
    t = t & "- For keyboard handling, prefer VK codes for comma/period (188/190) to avoid locale issues." & NL
    t = t & "- With Monday as week start, only Sunday is red and Saturday is blue." & NL

    BuildYearCalendarManual_EN = t
End Function

' ─────────────────────────────────────────────────────────────
' (선택) 강제로 영문/국문 도움말을 띄우고 싶을 때 호출할 헬퍼
' ─────────────────────────────────────────────────────────────
Public Sub ShowHelp_KO()
    On Error Resume Next
    frmYearCalHelp.Show vbModeless
    frmYearCalHelp.txtHelp.text = BuildYearCalendarManual_KO()
End Sub

Public Sub ShowHelp_EN()
    On Error Resume Next
    frmYearCalHelp.Show vbModeless
    frmYearCalHelp.txtHelp.text = BuildYearCalendarManual_EN()
End Sub

