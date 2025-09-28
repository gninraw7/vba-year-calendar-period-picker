VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYearCalendar 
   Caption         =   "Year Calendar (Period Picker)  ♣ Author :  gninraw7@naver.com"
   ClientHeight    =   9416.001
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   19140
   OleObjectBlob   =   "frmYearCalendar.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmYearCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit

' 연도 동기화 중 재귀 방지 플래그
Private mSyncingYear As Boolean

' === Week-start & Weekname style ===
Private Const REG_KEY_WEEK_START As String = "WeekStart"         ' "Sun" / "Mon"
Private Const REG_KEY_WEEKNAME_STYLE As String = "WeekNameStyle" ' "Short" / "Full"

Private Enum eWeekStart
    WeekSun = 0   ' 주 시작: 일요일
    WeekMon = 1   ' 주 시작: 월요일
End Enum

Private Enum eWeekNameStyle
    WkShort = 0   ' Sun, Mon, ...
    WkFull = 1    ' Sunday, Monday, ...
End Enum

Private mWeekStart As eWeekStart
Private mWeekNameStyle As eWeekNameStyle


' ===== 언어 상태 =====
Private Enum eLang
    LangK = 0   ' Korean
    LangE = 1   ' English
End Enum

Private mLang As eLang
Private mFmtDate As String         ' 예: "yyyy-mm-dd" 또는 "mmm d, yyyy"


' 기본값(최초 실행 시)
Private Const DEF_FMT_DATE_K As String = "yyyy-mm-dd"
Private Const DEF_FMT_DATE_E As String = "yyyy-mm-dd"     ' 필요 시 "mmm d, yyyy" 로 바꿔쓰면 됨
Private Const DEF_FMT_TITLE_K As String = "yyyy""년"" m""월"""
Private Const DEF_FMT_TITLE_E As String = "mmmm yyyy"

' ========= Registry I/O & 유틸 =========
Private Const REG_APP As String = "PeriodPicker"
Private Const REG_SEC As String = "Holidays"

' === Registry (Layout prefs) ===
Private Const REG_SEC_LAYOUT As String = "Layout"
Private Const REG_KEY_MODE As String = "LayoutMode"   ' "0"~"3"
Private Const REG_KEY_SLOT As String = "AnchorSlot"   ' "1"~"12"
Private Const REG_KEY_FROM_ONLY As String = "Apply_From_Only"   ' "1"."0"
Private Const REG_KEY_EXCLUDE_NONBIZ As String = "Exclude_Non_Biz"   ' "1"."0"
Private Const REG_KEY_SHOW_KEEP As String = "Show_Keep"   ' "1"."0"
Private Const REG_KEY_AUTO_NEXTROW As String = "AutoNextRow"


' === i18n / DateFormat ===
Private Const REG_SEC_I18N As String = "I18N"
Private Const REG_KEY_LANG As String = "Lang"     ' "K" / "E"

Private Const REG_SEC_FMT As String = "Format"
Private Const REG_KEY_FMT_DATE As String = "Date"        ' txtFromSel/ToSel, 반환, 셀 NumberFormat

' === Month title formats per language ===
Private Const REG_KEY_FMT_TITLE As String = "MonthTitle" ' 월 타이틀 라벨/시트 타이틀
Private Const REG_KEY_FMT_TITLE_K As String = "MonthTitle_K"
Private Const REG_KEY_FMT_TITLE_E As String = "MonthTitle_E"

Private mFmtMonthTitle As String      ' 현재 언어에 따라 선택되어 사용
Private mFmtMonthTitleK As String     ' 한글 모드 전용
Private mFmtMonthTitleE As String     ' 영어 모드 전용


' ==== 키 마스크 ====
Private Const SHIFT_MASK As Integer = 1
Private Const CTRL_MASK  As Integer = 2
Private Const ALT_MASK   As Integer = 4

' === 배치 모드 확장 ===
Private Enum eLayoutMode
    lmNormal = 0          ' 1월~12월
    lmCurrentFirst = 1    ' 현재월이 1행1열
    lmCurrentLast = 2     ' 현재월이 3행4열
    lmCurrentAtSlot = 3   ' ★현재월을 지정 슬롯(1..12)에 배치
End Enum

Private mLayoutMode As eLayoutMode

Private Const VK_OEM_PERIOD     As Long = 190
Private Const VK_OEM_COMMA      As Long = 188

' === 현재월 슬롯 앵커(1..12) ===
Private mAnchorSlot As Long  ' optCurrentAtSlot 모드에서만 사용

' === 월 박스 배치 행/열 (4x3) ===
Private Const GRID_COLS As Long = 4
Private Const GRID_ROWS As Long = 3

' === 추가: 현재월 배경 라벨들 ===
Private lblMonthBG(1 To 12) As MSForms.Label

' 색/여백
Private mCurMonthBG As Long
Private mCurMonthBorder As Long
Private mBGPad As Single

' ========= 색상 (Const 금지: RGB는 함수라 런타임에 설정) =========
Private mClrBG As Long
Private mClrSunFg As Long, mClrSatFg As Long, mClrWeekFg As Long
Private mClrHolBg As Long, mClrTodayBg As Long, mClrRangeBg As Long
Private mClrMonthTitle As Long
Private mClrCurMonthBg As Long   ' 현재월 강조 배경

' === Month label 연도 구분 색 ===
Private mYearOddBG  As Long   ' 홀수 연도 배경
Private mYearEvenBG As Long   ' 짝수 연도 배경
Private mYearFG     As Long   ' 공통 전경 (글자색)

' ========= 반환 타깃 =========
Private Enum ePushMode
    PushNone = 0
    PushToRange = 1          ' 시트 범위(Selection 또는 지정 Range)
    PushToTextBoxes = 2      ' 다른 폼의 TextBox 두 개(From/To)
End Enum

Private mPushMode As ePushMode
Private mTargetRange As Range
Private mTargetTextFrom As MSForms.TextBox
Private mTargetTextTo As MSForms.TextBox

' ========= 폰트/레이아웃 =========
Private Const FONT_NAME As String = "Segoe UI"
Private Const FONT_SIZE As Single = 8

Private Const MARGIN As Single = 2
Private Const GAP_MONTH_X As Single = 2
Private Const GAP_MONTH_Y As Single = 4
Private Const TITLE_H As Single = 18
Private Const WEEK_H As Single = 16
Private Const CELL_W As Single = 20
Private Const CELL_H As Single = 14
Private Const CELL_GAP As Single = 2

' ========= 상태 =========
Private mYear As Long
Private mCreated As Boolean

' 동적 컨트롤 배열
Private lblMonth(1 To 12) As MSForms.Label
Private lblWeek(1 To 12, 1 To 7) As MSForms.Label
Private tbDay(1 To 12, 1 To 6, 1 To 7) As MSForms.TextBox

' 이벤트 훅 보관
Private mHooks As Collection

' From/To 선택
Private mHasFrom As Boolean, mHasTo As Boolean
Private mFrom As Date, mTo As Date

' 공휴일 캐시
Private HolidayByDate As Object       ' key "yyyy-mm-dd" → name
Private HolidayYearLoaded As Object   ' key "2025" → True

'========= 2025.09.20 Task Panel =========
' Task Panel 토글 사이즈
Private mBaseWidth As Single
Private mTaskWidth As Single
Private mTaskVisible As Boolean

' 폼 내부 사용: 일자 -> TextBox 매핑 (오버레이용)
Private DayBoxByDate As Object ' Scripting.Dictionary
Private OrigBorderColor As Object     ' "yyyy-mm-dd" -> Long
Private OrigToolTip As Object         ' "yyyy-mm-dd" -> String

' 사용자 선호: Task 오버레이/선택연동
Private mTaskOverlay As Boolean
Private mTaskLinkSel As Boolean

'=== Task Color 매핑 ===
Private TaskColorByName As Object  ' key: TaskKey(보통 TaskName), val: Long(OLE Color)
Private Const REG_SEC_TASKCOLOR As String = "TaskColors"  ' PeriodPicker\TaskColors 밑에 저장
Private Const SEC_TASKS As String = "Tasks"      ' SaveSetting/GetSetting 섹션

'=== Task Panel prefs (Registry) ===
Private Const REG_SEC_TASKPANEL As String = "TaskPanel"
Private Const REG_KEY_TASK_VISIBLE As String = "Visible"   ' "1" / "0"
Private Const REG_KEY_TASK_OVERLAY As String = "Overlay"   ' "1" / "0"
Private Const REG_KEY_TASK_LINKSEL As String = "LinkSel"   ' "1" / "0"
Private Const REG_KEY_SHOWALL As String = "ShowAll"
Private Const REG_KEY_TASK_CATEGORY As String = "Category"   ' 마지막 선택 카테고리 저장

' 선택된 Task들의 개별 구간(여러 개)을 칠하기 위한 보관소
Private mSelRanges As Collection   ' 각 아이템: Variant(2) = Array(DateStart, DateEnd)

' === Overlay 배경 복원 및 겹침 강도 관리 ===
Private OrigBackColor As Object    ' key: "yyyy-mm-dd" -> Long(원래 BackColor)
Private OverlayHitCount As Object  ' key: "yyyy-mm-dd" -> Long(겹침 횟수)

' 선택 하이라이트 재진입 방지
Private mInSelOverlay As Boolean

Private mRefreshing As Boolean

Private Sub SetupColors()
    mClrBG = vbWhite
    mClrSunFg = RGB(204, 0, 0)
    mClrSatFg = RGB(0, 90, 200)
    mClrWeekFg = vbBlack
    mClrHolBg = RGB(255, 235, 235)
    mClrTodayBg = RGB(204, 255, 204)
    mClrRangeBg = RGB(255, 247, 204)
    mClrMonthTitle = vbBlack
    mClrCurMonthBg = RGB(240, 248, 255)  ' 은은한 하이라이트(AliceBlue 계열)
    
    ' 연도 구분 배경/글자색 (원하는 색으로 조정 가능)
    mYearOddBG = RGB(255, 248, 230)    ' 홀수년: 약간 웜톤
    mYearEvenBG = RGB(236, 244, 255)   ' 짝수년: 약간 쿨톤
    mYearFG = RGB(30, 30, 30)
    
    mCurMonthBG = RGB(255, 247, 205)     ' 은은한 노랑
    mCurMonthBorder = RGB(240, 170, 0)   ' 테두리 색
    mBGPad = 2                           ' 블록 바깥쪽 여백(px)
    
End Sub

Private Sub LoadLayoutPrefs()
    Dim sMode As String, sSlot As String
    Dim m As Long, slot As Long

    sMode = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_MODE, "0")
    sSlot = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_SLOT, "1")

    m = CLng(Val(sMode))
    If m < 0 Or m > 3 Then m = 0

    slot = CLng(Val(sSlot))
    If slot < 1 Or slot > 12 Then slot = 1

    ' 내부 상태 반영
    mLayoutMode = m
    mAnchorSlot = slot

    ' UI 반영
    optNormal.Value = (mLayoutMode = lmNormal)
    optCurrentFirst.Value = (mLayoutMode = lmCurrentFirst)
    optCurrentLast.Value = (mLayoutMode = lmCurrentLast)
    optCurrentAtSlot.Value = (mLayoutMode = lmCurrentAtSlot)

    txtAnchorSlot.text = CStr(mAnchorSlot)
    spnSlot.Value = mAnchorSlot
    EnableAnchorSlotUI (mLayoutMode = lmCurrentAtSlot)
    
    spnYear.Min = 1901
    spnYear.Max = 9998
    spnYear.Value = mYear
    
    Dim sFrom As String, sExcludeNonBiz As String, sShowKeep As String

    sFrom = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_FROM_ONLY, "0")
    If sFrom = "0" Then
        chkFromOnly.Value = False
    Else
        chkFromOnly.Value = True
    End If
    
    sExcludeNonBiz = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_EXCLUDE_NONBIZ, "1")
    If sExcludeNonBiz = "1" Then
        chkExcludeNonBiz.Value = True
    Else
        chkExcludeNonBiz.Value = False
    End If
    
    sShowKeep = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_SHOW_KEEP, "0")
    If sShowKeep = "0" Then
        chkShowKeep.Value = False
    Else
        chkShowKeep.Value = True
    End If
    
    ' Auto next row (apply 후 선택영역을 아래로 이동)
    Dim sAuto As String
    sAuto = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_AUTO_NEXTROW, "0")
    Me.chkAutoNextRow.Value = (sAuto = "1")
    
    
End Sub

' === 저장 ===
Private Sub SaveLayoutPrefs()
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_MODE, CStr(mLayoutMode)
    SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_SLOT, CStr(mAnchorSlot)
End Sub

Private Sub btnCurrYear_Click()
    SetBaseYear Year(Date)   ' 스핀값까지 함께 동기화
End Sub

Private Sub chkFromOnly_Click()
    On Error Resume Next
    If chkFromOnly.Value Then
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_FROM_ONLY, "1"
    Else
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_FROM_ONLY, "0"
    End If
End Sub

Private Sub chkShowKeep_Click()
    On Error Resume Next
    If chkShowKeep.Value Then
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_SHOW_KEEP, "1"
    Else
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_SHOW_KEEP, "0"
    End If
End Sub

Private Sub chkAutoNextRow_Click()
    On Error Resume Next
    If Me.chkAutoNextRow.Value Then
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_AUTO_NEXTROW, "1"
    Else
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_AUTO_NEXTROW, "0"
    End If
End Sub

Private Sub txtTaskFilter_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    btnTaskFilterApply_Click
End Sub

' ========= 초기화 =========
Private Sub UserForm_Initialize()
    SetupColors
    
    LoadAllHolidaysIntoGlobalSet

    mYear = Year(Date)
    txtBaseYear.text = CStr(mYear)
    
    ' 공휴일 사전
    Set HolidayByDate = CreateObject("Scripting.Dictionary")
    HolidayByDate.CompareMode = vbTextCompare
    Set HolidayYearLoaded = CreateObject("Scripting.Dictionary")
    
    Me.BackColor = vbWhite

    CreateAllMonthBlocks
    
    txtFromSel.BackColor = RGB(255, 244, 204)
    txtToSel.BackColor = RGB(255, 244, 204)
    
   
    ' === 레이아웃 선호도 로드 ===
    LoadLayoutPrefs
    
    LoadI18NAndFormats
    chkKE.Value = (mLang = LangE)
    
    ApplyStaticUIStrings
    
    RenderAllMonths
    
    ReapplyLanguageOnCalendar
    
    txtFromSel.text = "": txtToSel.text = ""
    txtFromSel.Locked = True: txtToSel.Locked = True
    
    btnClose.Cancel = True
    
    EnableMouseScroll Me
    
'    EnableMouseScroll Me, True, True, False          ' 기존처럼 딱 한 줄
'    WheelBridge.RegisterWheelSink Me   ' ★ 추가: 이 폼을 휠 수신자로 등록
    
    SetBaseYear mYear, False
    
    InitTaskPanel
    
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
'    WheelBridge.UnregisterWheelSink Me ' ★ 추가: 폼 닫힐 때 등록 해제
End Sub

' 12칸 각각에 표시할 "그 달의 1일" 날짜를 채워줌
Private Sub BuildMonthMap(ByVal baseYear As Long, ByVal curMonth As Long, _
                          ByVal mode As eLayoutMode, ByRef slotDate() As Date, _
                          Optional ByVal anchorSlot As Long = 1)
    ReDim slotDate(1 To 12)

    Dim startD As Date
    Select Case mode
        Case lmNormal
            ' 1월~12월 (모두 baseYear)
            startD = DateSerial(baseYear, 1, 1)

        Case lmCurrentFirst
            ' 슬롯1 = baseYear의 현재월 → 이후 11개월
            startD = DateSerial(baseYear, curMonth, 1)

        Case lmCurrentLast
            ' 슬롯12 = baseYear의 현재월 → 앞쪽 11개월
            startD = DateAdd("m", -11, DateSerial(baseYear, curMonth, 1))

        Case lmCurrentAtSlot
            ' 슬롯(anchorSlot) = baseYear의 현재월
            If anchorSlot < 1 Or anchorSlot > 12 Then anchorSlot = 1
            ' 슬롯1이 되도록 현재월을 (-(anchorSlot-1))만큼 당긴 것이 시작점
            startD = DateAdd("m", -(anchorSlot - 1), DateSerial(baseYear, curMonth, 1))
    End Select

    Dim i As Long
    For i = 1 To 12
        slotDate(i) = DateAdd("m", i - 1, startD)
    Next
End Sub

Private Sub chkExcludeNonBiz_Click()
    PaintRange
    On Error Resume Next
    If chkExcludeNonBiz.Value Then
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_EXCLUDE_NONBIZ, "1"
    Else
        SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_EXCLUDE_NONBIZ, "0"
    End If
    
    RefreshTaskOverlayIfOn
    ' ★ 선택 하이라이트도 옵션 반영해서 다시 덮어 칠
    ApplySelectedRangeOverlay
End Sub

Private Sub RenderAllMonths()
    Dim baseYear As Long: baseYear = mYear
    
    Dim curM As Long: curM = Month(Date)  ' 시스템 현재월 기준 앵커

    Dim mapDate() As Date
    
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot

    Dim i As Long
    For i = 1 To 12
        Dim y As Long, m As Long
        y = Year(mapDate(i))
        m = Month(mapDate(i))
        RenderMonthByBlock i, y, m    ' 내부에서 DrawMonthCells 호출
    Next
    
    EnsureMonthBGCreated
    SizeAllMonthBG
    UpdateCurrentMonthBG mapDate   ' ← 현재월 블록만 표시
    
    BuildDayBoxMapFromGrid
    ClearAllDayBoxOverlay
    
    Dim t As Collection
    Set t = TasksFromListBox(Me.lstTask)
    
    EnsureColorsForTasks t

    If mTaskOverlay Then
        ApplyTaskOverlay t
    Else
        ClearTaskOverlay
    End If
    
    ' 선택 구간은 최후에 다시 덮어 칠하기
    ApplySelectedRangeOverlay
    
End Sub

' mBlock: 1..12 (화면상의 위치), y/m: 실제 달력의 년/월
Private Sub RenderMonthByBlock(ByVal mBlock As Long, ByVal y As Long, ByVal m As Long)
    DrawMonthCells mBlock, y, m
End Sub

' mBlock : 화면상의 월 블록 인덱스(1..12)
' y, m   : 실제 표시할 연/월
Private Sub DrawMonthCells(ByVal mBlock As Long, ByVal y As Long, ByVal m As Long)
    Dim r As Long, c As Long, d As Long
    Dim firstDay As Date, lastDay As Long, startCol As Long
    Dim tb As MSForms.TextBox, dT As Date
    Dim tip As String, holName As String

    ' ==== 색상(상수 대신 런타임 할당) ====
    Dim CLR_BG As Long:            CLR_BG = RGB(255, 255, 255)
    Dim CLR_SUN_BG As Long:        CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:        CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_TODAY_BG As Long:      CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_HOLI_BG As Long:       CLR_HOLI_BG = RGB(255, 235, 235)

    Dim CLR_TEXT As Long:          CLR_TEXT = vbBlack
    Dim CLR_SUN_FG As Long:        CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:        CLR_SAT_FG = RGB(0, 90, 200)

    ' ==== 타이틀 (yyyy년mm월) ====
    lblMonth(mBlock).Caption = MonthTitleOut(y, m)
    
    ' ==== 타이틀 배경: 연도 짝/홀수 구분 ====
    With lblMonth(mBlock)
        .BackStyle = fmBackStyleOpaque
        If (y Mod 2) = 0 Then
            .BackColor = mYearEvenBG   ' 짝수년
        Else
            .BackColor = mYearOddBG    ' 홀수년
        End If
        .ForeColor = mYearFG
    End With

    ' ==== 모든 셀 초기화 ====
    For r = 1 To 6
        For c = 1 To 7
            Set tb = tbDay(mBlock, r, c)
            With tb
                .text = ""
                .Tag = ""
                .ControlTipText = ""
                .BackColor = CLR_BG
                .ForeColor = CLR_TEXT
                .Font.Bold = False
                .Font.name = FONT_NAME
                .Font.Size = FONT_SIZE
                .TextAlign = fmTextAlignCenter
                .Locked = True           ' 편집 방지(클릭 이벤트는 수신)
                .BorderStyle = 0   ' ★ 기본은 무테
            End With
        Next
    Next

    ' ==== 월 배치 계산 ====
    firstDay = DateSerial(y, m, 1)
    lastDay = Day(DateSerial(y, m + 1, 0))              ' 말일
    startCol = Weekday(firstDay, FirstDOWParam())  ' Mon 시작이면 vbMonday 기준

    r = 1: c = startCol

    ' ==== 날짜 채우기 ====
    For d = 1 To lastDay
        Set tb = tbDay(mBlock, r, c)
        dT = DateSerial(y, m, d)

        ' 기본 값
        tb.text = CStr(d)
        tb.Tag = Format$(dT, "yyyy-mm-dd")
        tb.ControlTipText = FmtDateOut(dT) ' 언어 서식 반영

        ' 주말 전경색/배경색 (날짜 기준)
        If IsSundayDate(dT) Then
            tb.ForeColor = CLR_SUN_FG
            tb.BackColor = CLR_SUN_BG
        ElseIf IsSaturdayDate(dT) Then
            tb.ForeColor = CLR_SAT_FG
            tb.BackColor = CLR_SAT_BG
        Else
            tb.ForeColor = CLR_TEXT
        End If

        ' 공휴일(이름 조회) → 배경/툴팁 보강
        holName = GetHolidayNameIfAny(dT)
        If Len(holName) > 0 Then
            tb.BackColor = CLR_HOLI_BG
            tb.ForeColor = CLR_SUN_FG         ' 관례적으로 공휴일은 적색 계열
            tb.ControlTipText = tb.ControlTipText & vbCrLf & "[공휴일] " & holName
        End If

        ' 오늘 강조(배경/볼드) - 공휴일보다 우선 표시하고 싶으면 순서 조정
        If dT = Date Then
            tb.BackColor = CLR_TODAY_BG
            tb.Font.Bold = True
        End If

        ' 다음 셀로 이동
        c = c + 1
        If c > 7 Then c = 1: r = r + 1: If r > 6 Then Exit For
    Next

    ' 선택범위 덮어칠 것 있으면 호출
    On Error Resume Next
    PaintRange
    On Error GoTo 0
End Sub

Private Function GetHolidayNameIfAny(ByVal d As Date) As String
    On Error Resume Next
    If gHolidaySet Is Nothing Then Exit Function
    Dim k As Date: k = DateSerial(Year(d), Month(d), Day(d)) ' 시간 0:00 정규화
    If gHolidaySet.Exists(k) Then GetHolidayNameIfAny = CStr(gHolidaySet(k))
End Function

Public Sub HandleDayMouse(ByVal tb As MSForms.TextBox, ByVal Button As Integer)
    Dim s As String, d As Date
    s = Trim$(tb.Tag)
    If Len(s) = 0 Then Exit Sub                        ' 빈칸 셀
    If Not TryParseYMD(s, d) Then Exit Sub             ' 방어
    
    If Button = 1 Then
        ' From
        mFrom = d: mHasFrom = True
    ElseIf Button = 2 Then
        ' To
        mTo = d: mHasTo = True
    Else
        Exit Sub
    End If

    ' From/To 정규화(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' 상단 표시
    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")

    PaintRange        ' 선택 구간 다시 칠하기
    UpdateRangeInfo   ' "Business Day / 총일수" 갱신
    ApplySelectedRangeOverlay
    'RefreshTaskListAndOverlay
End Sub

' 필요 시 한 번만 생성
Private Sub EnsureMonthBGCreated()
    Dim i As Long
    For i = 1 To 12
        If lblMonthBG(i) Is Nothing Then
            Set lblMonthBG(i) = Me.Controls.add("Forms.Label.1", "lblMonthBG_" & CStr(i), True)
            With lblMonthBG(i)
                .Caption = ""
                .BackStyle = fmBackStyleOpaque
                .BackColor = mCurMonthBG
                On Error Resume Next
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = mCurMonthBorder
                On Error GoTo 0
                .visible = False
                .ZOrder 1      ' 뒤로 보내기 (fmZOrderBack)
            End With
        End If
    Next
End Sub

' 월 블록 1..12의 사각형을 계산하여 배경 라벨 크기 맞춤
Private Sub SizeAllMonthBG()
    Dim i As Long
    For i = 1 To 12
        SizeOneMonthBG i
    Next
End Sub

Private Sub SizeOneMonthBG(ByVal i As Long)
    On Error Resume Next
    If lblMonth(i) Is Nothing Then Exit Sub

    ' 블록의 좌상단: 타이틀 라벨
    Dim leftEdge As Single, topEdge As Single
    leftEdge = lblMonth(i).Left
    topEdge = lblMonth(i).Top

    ' 블록의 우하단: 6행 7열 TextBox (존재 가정)
    Dim br As MSForms.TextBox
    Set br = tbDay(i, 6, 7)

    If br Is Nothing Then Exit Sub

    Dim rightEdge As Single, bottomEdge As Single
    rightEdge = br.Left + br.Width
    bottomEdge = br.Top + br.Height

    With lblMonthBG(i)
        .Left = leftEdge - mBGPad
        .Top = topEdge - mBGPad
        .Width = (rightEdge - leftEdge) + 2 * mBGPad + 1
        .Height = (bottomEdge - topEdge) + 2 * mBGPad + 1
        .ZOrder 1   ' 뒤로
    End With
End Sub

' 현재월(시스템 오늘 기준)이 화면에 있으면 해당 블록만 보이게
Private Sub UpdateCurrentMonthBG(ByRef mapDate() As Date)
    Dim i As Long, yT As Long, mT As Long
    yT = Year(Date): mT = Month(Date)

    ' 모두 숨김
    For i = 1 To 12
        If Not lblMonthBG(i) Is Nothing Then lblMonthBG(i).visible = False
    Next

    ' mapDate(i) = 각 슬롯의 "그 달의 1일"
    For i = 1 To 12
        If Year(mapDate(i)) = yT And Month(mapDate(i)) = mT Then
            If Not lblMonthBG(i) Is Nothing Then
                lblMonthBG(i).visible = True
                lblMonthBG(i).ZOrder 1   ' 안전하게 뒤로
            End If
            Exit For
        End If
    Next
End Sub

Private Sub optNormal_Click()
    mLayoutMode = lmNormal
    EnableAnchorSlotUI False
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' ★ 추가
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' ★ 추가
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentFirst_Click()
    mLayoutMode = lmCurrentFirst
    EnableAnchorSlotUI False
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' ★ 추가
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' ★ 추가
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentLast_Click()
    mLayoutMode = lmCurrentLast
    EnableAnchorSlotUI False
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' ★ 추가
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' ★ 추가
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentAtSlot_Click()
    mLayoutMode = lmCurrentAtSlot
    EnableAnchorSlotUI True
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' ★ 추가
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' ★ 추가
    RefreshTaskListAndOverlay
End Sub

Private Sub spnSlot_Change()
    Dim v As Long
    v = spnSlot.Value
    If v < 1 Then v = 1
    If v > 12 Then v = 12
    If v <> mAnchorSlot Then
        mAnchorSlot = v
        txtAnchorSlot.text = CStr(v)
        If mLayoutMode = lmCurrentAtSlot Then
            RenderAllMonths
            Call RefreshTaskOverlayIfOn   ' ★ 추가
            PaintRange
        End If
        SaveLayoutPrefs
        'LoadTasksForVisibleRangeAndOverlay   ' ★ 추가
        RefreshTaskListAndOverlay
    End If
End Sub

Private Sub RefreshTaskOverlayIfOn()
    ' 렌더 뒤에 매핑이 최신 상태임을 전제로, On이면 적용/Off면 해제
    BuildDayBoxMapFromGrid
    Dim t As Collection
    Set t = TasksFromListBox(Me.lstTask)
    If mTaskOverlay Then
        ApplyTaskOverlay t
    Else
        ClearTaskOverlay
    End If
End Sub

Private Sub EnableAnchorSlotUI(ByVal onoff As Boolean)
    txtAnchorSlot.Enabled = onoff
    spnSlot.Enabled = onoff
End Sub

Private Sub spnYear_Change()
    If mSyncingYear Then Exit Sub
    SetBaseYear spnYear.Value
    RefreshTaskListAndOverlay
End Sub

' 선택 초기화가 필요할 때 사용(옵션)
Private Sub ClearSelectionUI()
    mHasFrom = False
    mHasTo = False
    txtFromSel.text = ""
    txtToSel.text = ""
    PaintRange
    UpdateRangeInfo
    RefreshTaskListAndOverlay
End Sub

Private Sub txtAnchorSlot_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim v As Long
    v = CLng(Val(Trim$(txtAnchorSlot.text)))
    If v < 1 Then v = 1
    If v > 12 Then v = 12
    txtAnchorSlot.text = CStr(v)
    If v <> mAnchorSlot Then
        mAnchorSlot = v
        spnSlot.Value = v
        If mLayoutMode = lmCurrentAtSlot Then
            RenderAllMonths
            PaintRange
        End If
        SaveLayoutPrefs
    End If
End Sub


' ========= 동적 생성 =========
Private Sub CreateAllMonthBlocks()
    If mCreated Then Exit Sub
    Set mHooks = New Collection

    Dim m As Long, r As Long, c As Long
    For m = 1 To 12
        Dim rowIdx As Long, colIdx As Long
        ' m = 1..12
        rowIdx = (m - 1) \ GRID_COLS          ' 0..(GRID_ROWS-1)
        colIdx = (m - 1) Mod GRID_COLS        ' 0..(GRID_COLS-1)

        Dim originX As Single, originY As Single
        originX = MARGIN + colIdx * (7 * (CELL_W + CELL_GAP) + GAP_MONTH_X + 15) + 10
        originY = 48 + rowIdx * (TITLE_H + WEEK_H + 6 * (CELL_H + CELL_GAP) + GAP_MONTH_Y) + 10

        ' === ★ 현재월 배경 라벨(먼저 생성해서 뒤로 배치됨) ===
        Dim bgW As Single, bgH As Single
        bgW = 7 * (CELL_W + CELL_GAP)                  ' 월 전체 폭
        bgH = TITLE_H + WEEK_H + 6 * (CELL_H + CELL_GAP) - CELL_GAP  ' 마지막 간격 보정
        Set lblMonthBG(m) = AddBGLabel("lblMonthBG" & m, originX - 4, originY - 4, bgW + 8, bgH + 8)

        ' === 월 타이틀 ===
        Set lblMonth(m) = AddLabel("lblMonth" & m, originX, originY, 7 * (CELL_W + CELL_GAP) - 1, TITLE_H - 2, CStr(m) & "월", True)
        lblMonth(m).Font.name = "Segoe UI Semibold"
        lblMonth(m).Font.Size = FONT_SIZE + 2
        lblMonth(m).Font.Bold = True
        lblMonth(m).ForeColor = mClrMonthTitle

        ' === 요일 ===
        Dim cidx As Long
        For cidx = 1 To 7
            Set lblWeek(m, cidx) = AddLabel("lblWeek" & m & "_" & cidx, _
                            originX + (cidx - 1) * (CELL_W + CELL_GAP), originY + TITLE_H, _
                            CELL_W, WEEK_H - 1, WeekNameByPos(cidx), True)
            lblWeek(m, cidx).BorderStyle = 0
            
            ' 요일색(일/토만 강조) - 실제 요일 기준
            Dim realDow As Long: realDow = DowByPos(cidx)  ' 1=Sun .. 7=Sat
            If realDow = vbSunday Then
                lblWeek(m, cidx).ForeColor = mClrSunFg
            ElseIf realDow = vbSaturday Then
                lblWeek(m, cidx).ForeColor = mClrSatFg
            Else
                lblWeek(m, cidx).ForeColor = mClrWeekFg
            End If
        Next
        
        ' === 일자(TextBox) 6×7 ===
        For r = 1 To 6
            For c = 1 To 7
                Dim nm As String: nm = "tbD_" & m & "_" & r & "_" & c
                Set tbDay(m, r, c) = AddDayTextBox(nm, _
                    originX + (c - 1) * (CELL_W + CELL_GAP), _
                    originY + TITLE_H + WEEK_H + (r - 1) * (CELL_H + CELL_GAP), _
                    CELL_W, CELL_H)
                tbDay(m, r, c).BorderStyle = 0
            Next
        Next
    Next
    
    ' Control의 Event Hook
    Dim iCtrl As MSForms.control
    For Each iCtrl In Me.Controls
        If iCtrl.Tag <> "NOHOOK" Then
            HookOneControl iCtrl
        End If
    Next

    EnsureMonthBGCreated
    SizeAllMonthBG

    mCreated = True
    
End Sub

Private Sub HookOneControl(ByVal ctrl As Object)
    Dim hk As cDayBox
    Select Case TypeName(ctrl)
        Case "TextBox", "ComboBox", "ListBox", "CommandButton", "OptionButton", "SpinButton", "Label", "CheckBox"
            Set hk = New cDayBox
            hk.Hook ctrl, Me
            mHooks.add hk
    End Select
End Sub


Private Function AddBGLabel(ByVal nm As String, ByVal X As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim L As MSForms.Label
    Set L = Me.Controls.add("Forms.Label.1", nm, True)
    With L
        .Left = X: .Top = y: .Width = w: .Height = h
        .Caption = ""
        .BackStyle = fmBackStyleOpaque
        .BackColor = mClrCurMonthBg
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .visible = False       ' 기본은 숨김 → 현재월만 보이게
        '.TabStop = False
    End With
    Set AddBGLabel = L
End Function

Private Function AddLabel(ByVal nm As String, ByVal X As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single, ByVal txt As String, _
                          Optional ByVal center As Boolean = False) As MSForms.Label
    Dim L As MSForms.Label
    Set L = Me.Controls.add("Forms.Label.1", nm, True)
    With L
        .Left = X: .Top = y: .Width = w: .Height = h
        .Caption = txt
        .BackStyle = fmBackStyleOpaque
        .BackColor = mClrBG
        .Font.name = FONT_NAME
        .Font.Size = FONT_SIZE
        .TextAlign = IIf(center, fmTextAlignCenter, fmTextAlignLeft)
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .WordWrap = False
        '.TabStop = False
    End With
    Set AddLabel = L
End Function

Private Function AddDayTextBox(ByVal nm As String, ByVal X As Single, ByVal y As Single, _
                               ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.add("Forms.TextBox.1", nm, True)
    With t
        .Left = X: .Top = y: .Width = w: .Height = h
        .text = ""
        .BackColor = mClrBG
        .BorderStyle = 0
        .SpecialEffect = fmSpecialEffectFlat
        .Font.name = FONT_NAME
        .Font.Size = FONT_SIZE
        .TextAlign = fmTextAlignCenter
        .MultiLine = False
        .AutoTab = False
        '.TabStop = False
        .EnterKeyBehavior = False
        .Locked = True            ' 편집 방지(클릭은 가능)
        .Enabled = True           ' 마우스 이벤트를 위해 Enabled 유지
    End With
    Set AddDayTextBox = t
End Function

' ========= 렌더링 =========
Public Sub PaintRange()
    Dim have As Boolean, s As Date, e As Date
    Call GetSelectedRange(have, s, e)

    Dim mBlock As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox, sTag As String, dT As Date
    Dim excludeNonBiz As Boolean: excludeNonBiz = chkExcludeNonBiz.Value

    For mBlock = 1 To 12
        For r = 1 To 6
            For c = 1 To 7
                Set tb = tbDay(mBlock, r, c)
                sTag = Trim$(tb.Tag)

                ' 기본 복구
                ApplyBaseStyle tb, c

                ' 선택 구간은 가장 마지막에 강제
                If have Then
                    If Len(sTag) > 0 And TryParseYMD(sTag, dT) Then
                        If dT >= s And dT <= e Then
                            If excludeNonBiz Then
                                If Not (IsWeekend(dT) Or IsHoliday(dT)) Then
                                    PaintSelectedCell tb
                                End If
                            Else
                                PaintSelectedCell tb
                            End If
                        End If
                    End If
                End If
            Next
        Next
    Next

    UpdateRangeInfo

    ' ★ 항상 오버레이 재적용 (선택 구간은 위에서 가드되어 안전)
    RefreshTaskOverlayIfOn
End Sub


' 내부 상태에서만 From~To 구간을 계산
Private Sub GetSelectedRange(ByRef have As Boolean, ByRef s As Date, ByRef e As Date)
    have = False
    If Not (mHasFrom And mHasTo) Then Exit Sub
    If mTo < mFrom Then
        s = mTo: e = mFrom
    Else
        s = mFrom: e = mTo
    End If
    have = True
End Sub


' 열(colIndex) 기준의 기본 스타일 복구 후,
' 유효 날짜(=Tag가 yyyy-mm-dd)일 때만 주말/공휴일/오늘 색을 적용
Private Sub ApplyBaseStyle(ByVal tb As MSForms.TextBox, ByVal colIndex As Long)
    ' 기본 팔레트
    Dim CLR_BG As Long:       CLR_BG = vbWhite
    Dim CLR_TEXT As Long:     CLR_TEXT = vbBlack
    Dim CLR_SUN_BG As Long:   CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:   CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_TODAY_BG As Long: CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_HOLI_BG As Long:  CLR_HOLI_BG = RGB(255, 235, 235)
    Dim CLR_SUN_FG As Long:   CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:   CLR_SAT_FG = RGB(0, 90, 200)

    ' 1) 기본 초기화
    tb.Font.Bold = False
    tb.ForeColor = CLR_TEXT
    tb.BackColor = CLR_BG

    ' 2) 날짜 없는 칸(=Tag 없음/무효)은 여기서 끝 → 항상 흰 배경 유지
    Dim s As String: s = Trim$(tb.Tag)
    If Len(s) = 0 Then Exit Sub

    Dim d As Date
    If Not TryParseYMD(s, d) Then Exit Sub

    ' 3) 유효 날짜일 때만 주말/평일 색
    If IsSundayDate(d) Then
        tb.ForeColor = CLR_SUN_FG
        tb.BackColor = CLR_SUN_BG
    ElseIf IsSaturdayDate(d) Then
        tb.ForeColor = CLR_SAT_FG
        tb.BackColor = CLR_SAT_BG
    End If

    ' 4) 공휴일 적용(이름 있으면)
    Dim nm As String: nm = GetHolidayNameIfAny(d)
    If Len(nm) > 0 Then
        tb.BackColor = CLR_HOLI_BG
        tb.ForeColor = CLR_SUN_FG
    End If

    ' 5) 오늘 강조
    If d = Date Then
        tb.BackColor = CLR_TODAY_BG
        tb.Font.Bold = True
    End If
End Sub



' ========= UI 이벤트 =========
Private Sub txtBaseYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH
    Dim y As Long: y = CLng(Trim$(txtBaseYear.text))
    If y < 1901 Or y > 9998 Then GoTo EH
    SetBaseYear y                     ' 중앙 함수로 일원화
    Exit Sub
EH:
    MsgBox "기준년도를 정확히 입력하세요. 예) 2025", vbExclamation
    txtBaseYear.text = CStr(mYear)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ========= 공휴일 로딩/판정 =========
Private Sub EnsureHolidayYearLoaded(ByVal y As Long)
    Dim ky As String: ky = CStr(y)
    If HolidayYearLoaded.Exists(ky) Then Exit Sub

    Dim raw As String: raw = ReadHolidaysRaw(y) ' "yyyy-mm-dd|휴일명" 줄들
    If Len(raw) > 0 Then
        Dim ln() As String: ln = Split(raw, vbCrLf)
        Dim i As Long
        For i = LBound(ln) To UBound(ln)
            Dim s As String: s = Trim$(ln(i))
            If Len(s) = 0 Then GoTo ContinueNext
            Dim a() As String: a = Split(s, "|")
            Dim d As String: d = NormalizeYMD(a(0))
            If Len(d) = 0 Then GoTo ContinueNext
            HolidayByDate(d) = IIf(UBound(a) >= 1, a(1), "")
ContinueNext:
        Next
    End If

    HolidayYearLoaded.add ky, True
End Sub

Private Function ReadHolidaysRaw(ByVal y As Long) As String
    ReadHolidaysRaw = GetSetting(REG_APP, REG_SEC, CStr(y), "")
End Function

'----------------------

Public Sub SetTargetRange(ByVal rng As Range)
    Set mTargetRange = rng
    mPushMode = PushToRange
End Sub

Public Sub SetTargetTextBoxes(ByVal tbFrom As MSForms.TextBox, ByVal tbTo As MSForms.TextBox)
    Set mTargetTextFrom = tbFrom
    Set mTargetTextTo = tbTo
    mPushMode = PushToTextBoxes
End Sub

' 더블클릭: 미설정된 끝점을 채우고 바로 반환
Public Sub HandleDayDbl(ByVal tb As MSForms.TextBox)
    If Len(tb.Tag) = 0 Then Exit Sub
    Dim d As Date: d = CDate(Split(CStr(tb.Tag), "|")(0))

    If Not mHasFrom And Not mHasTo Then
        mFrom = d: mTo = d: mHasFrom = True: mHasTo = True
    ElseIf mHasFrom And Not mHasTo Then
        mTo = d: mHasTo = True
    ElseIf Not mHasFrom And mHasTo Then
        mFrom = d: mHasFrom = True
    Else
        ' 둘 다 있으면 범위는 그대로 두고 즉시 반환
    End If

    ' From<=To 보정
    Dim t As Date
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then t = mFrom: mFrom = mTo: mTo = t
    End If

    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")

    PaintRange
    PushRangeNow
End Sub

' 반환 실행(버튼/더블클릭 공용)
Private Sub PushRangeNow()
    If Not mHasFrom And Not mHasTo Then
        MsgBox "선택된 날짜가 없습니다.", vbExclamation: Exit Sub
    End If
    'If mHasFrom And Not mHasTo Then mTo = mFrom: mHasTo = True
    'If mHasTo And Not mHasFrom Then mFrom = mTo: mHasFrom = True
    If mHasFrom And mHasTo Then
        Dim t As Date
        If mTo < mFrom Then
            t = mFrom
            mFrom = mTo
            mTo = t
        End If
    End If

    Dim sFrom As String, sTo As String

    sFrom = IIf(mHasFrom, FmtDateOut(mFrom), "")
    sTo = IIf(mHasTo, FmtDateOut(mTo), "")

    Select Case mPushMode
        Case PushToTextBoxes
            If Not mTargetTextFrom Is Nothing Then mTargetTextFrom.text = sFrom
            If Not mTargetTextTo Is Nothing Then mTargetTextTo.text = sTo
            Unload Me

        Case PushToRange
            Dim rng As Range
            
'            If chkShowKeep = False Then
'                If mTargetRange Is Nothing Then
'                    On Error Resume Next
'                    Set rng = Selection
'                    On Error GoTo 0
'                Else
'                    Set rng = mTargetRange
'                End If
'                If rng Is Nothing Then
'                    MsgBox "대상 범위가 없습니다. 셀을 선택하거나 SetTargetRange로 지정하세요.", vbExclamation
'                    Exit Sub
'                End If
'            Else
                Set rng = Selection
                On Error GoTo 0
                If rng Is Nothing Then
                    MsgBox "대상 범위가 없습니다. 셀을 선택 하세요.", vbExclamation
                    Exit Sub
                End If
'            End If
            
            If chkFromOnly Then
                If Not mHasFrom Then
                    MsgBox "선택된 날짜가 없습니다.", vbExclamation: Exit Sub
                Else
                    ApplyDateToCell rng.Cells(1, 1), mFrom
                End If
            Else
                If rng.Cells.Count >= 2 Then
                    If mHasFrom Then
                        ApplyDateToCell rng.Cells(1, 1), mFrom
                    End If
                    If mHasTo Then
                        ApplyDateToCell rng.Cells(1, 2), mTo
                    End If
                Else
                    If mHasFrom Then
                        ApplyDateToCell rng.Cells(1, 1), mFrom
                    End If
                    If mHasTo Then
                        ApplyDateToCell rng.Cells(1, 1).Offset(0, 1), mTo
                    End If
                End If
            End If
            
            If Me.chkAutoNextRow.Value Then
                MoveSelectionDownOneRowOptional
            End If
            
            If Not chkShowKeep Then
                Unload Me
            End If

        Case Else
            ' 기본: Selection에 씀
            Dim tgt As Range
            On Error Resume Next
            Set tgt = Selection
            On Error GoTo 0
            If tgt Is Nothing Then
                Me.Tag = sFrom & "|" & sTo
                Me.Hide
            Else
                If tgt.Cells.Count >= 2 Then
                    If mHasFrom Then
                        ApplyDateToCell tgt.Cells(1, 1), mFrom
                    End If
                    If mHasTo Then
                        ApplyDateToCell tgt.Cells(1, 2), mTo
                    End If
                Else
                    If mHasFrom Then
                        ApplyDateToCell tgt.Cells(1, 1), mFrom
                    End If
                    If mHasTo Then
                        ApplyDateToCell tgt.Cells(1, 1).Offset(0, 1), mTo
                    End If
                End If
                
                If Me.chkAutoNextRow.Value Then
                    MoveSelectionDownOneRowOptional
                End If
                
                Unload Me
            End If
    End Select
End Sub

' 현재 선택 영역을 아래로 1행 이동(크기 유지)
Private Sub MoveSelectionDownOneRowOptional()
    On Error Resume Next
    Dim sel As Range
    Set sel = Selection
    If sel Is Nothing Then Exit Sub

    Dim nextTop As Long
    nextTop = sel.row + 1
    If nextTop <= sel.Worksheet.Rows.Count Then
        sel.Offset(1, 0).Resize(sel.Rows.Count, sel.Columns.Count).Select
    End If
End Sub


' 적용(반환)
Private Sub btnApplyRange_Click()
    PushRangeNow
End Sub

' Clear: From/To 초기화 + 칠한 색 되돌리기
Private Sub btnClear_Click()
    ClearSelectionUI
End Sub

' ==== 공휴일 여부 (레지스트리 로딩해 둔 gHolidaySet 사용) ====
Private Function IsHoliday(ByVal d As Date) As Boolean
    On Error Resume Next
    If gHolidaySet Is Nothing Then Exit Function
    ' 자정으로 정규화(시간이 0:00 이 아니어도 안전)
    IsHoliday = gHolidaySet.Exists(CDate(Int(CDbl(d))))
End Function

' ==== Business Day 개수(양 끝 포함) ====
Public Function CountBusinessDays(ByVal dFrom As Date, ByVal dTo As Date) As Long
    Dim s As Date, e As Date, d As Date, n As Long
    If dTo < dFrom Then s = dTo: e = dFrom Else s = dFrom: e = dTo
    For d = s To e
        If Not IsWeekend(d) Then
            If Not IsHoliday(d) Then n = n + 1
        End If
    Next
    CountBusinessDays = n
End Function

' ==== 표시 업데이트: "Business Day / 총일수" ====
Public Sub UpdateRangeInfo()
    If Not (mHasFrom And mHasTo) Then
        txtRangeInfo.text = "": Exit Sub
    End If

    Dim s As Date, e As Date
    If mTo < mFrom Then s = mTo: e = mFrom Else s = mFrom: e = mTo

    Dim totalDays As Long, biz As Long
    totalDays = CLng(e - s) + 1
    biz = CountBusinessDays(s, e)

    txtRangeInfo.text = CStr(biz) & " / " & CStr(totalDays)
End Sub

' 널/에러 안전 문자열 변환
Private Function NzCStr(ByVal v As Variant) As String
    If IsError(v) Or isNull(v) Or IsEmpty(v) Then NzCStr = "" Else NzCStr = CStr(v)
End Function

Private Sub btnExport_Click()
    ExportCalendarCurrentLayout
End Sub

' 현재 레이아웃 그대로 새 시트에 출력
Private Sub ExportCalendarCurrentLayout()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As String
    nm = "YearCal_" & Format(Now, "yymmdd_hhnnss")
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
    On Error Resume Next
    ws.name = nm
    On Error GoTo 0
    
    ' 팔레트(런타임 변수)
    Dim CLR_TEXT As Long:     CLR_TEXT = vbBlack
    Dim CLR_BG As Long:       CLR_BG = vbWhite
    Dim CLR_GRID As Long:     CLR_GRID = RGB(210, 210, 210)
    Dim CLR_SUN_FG As Long:   CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:   CLR_SAT_FG = RGB(0, 90, 200)
    Dim CLR_SUN_BG As Long:   CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:   CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_HOLI_BG As Long:  CLR_HOLI_BG = RGB(255, 235, 235)
    Dim CLR_TODAY_BG As Long: CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_SEL_BG As Long:   CLR_SEL_BG = RGB(33, 92, 152) ' From~To 범위  CLR_SEL_BG = RGB(255, 244, 204) ' From~To 범위
    Dim CLR_SEL_FG As Long:   CLR_SEL_FG = RGB(255, 255, 255) ' From~To 범위
    
    Dim YearOddBG As Long:    YearOddBG = RGB(255, 248, 230)
    Dim YearEvenBG As Long:   YearEvenBG = RGB(236, 244, 255)
    Dim YearFG As Long:       YearFG = RGB(30, 30, 30)
    
    ' 배치 메트릭스
    Const COL0 As Long = 2         ' 첫 블록 좌상단 컬럼
    Const ROW0 As Long = 4         ' 첫 블록 좌상단 행 (1~3행은 헤더/요약 용)
    Const COLS_PER_MONTH As Long = 7  ' 요일 7칸
    Const ROWS_TITLE As Long = 1
    Const ROWS_WEEK As Long = 1
    Const ROWS_DAYS As Long = 6
    Const ROWS_PER_MONTH As Long = ROWS_TITLE + ROWS_WEEK + ROWS_DAYS
    Const GAP_COLS As Long = 2
    Const GAP_ROWS As Long = 1
    
    'Dim weekHdr: weekHdr = Split("일,월,화,수,목,금,토", ",")
    
    ' 기준연도/현재월 & 맵(12칸)
    Dim baseYear As Long: baseYear = CLng(Val(txtBaseYear.text))
    Dim curM As Long: curM = Month(Date)
    Dim mapDate() As Date
    
    Application.ScreenUpdating = False
    
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot   ' (앵커 슬롯 모드 지원)
    
    BuildDayBoxMapFromGrid    ' ★ UI에 보이는 DayBox 매핑 최신화
    
    Dim have As Boolean, sRange As Date, eRange As Date
    Call GetSelectedRange(have, sRange, eRange)
    
    ' 상단 요약(레이아웃/선택)
    ws.Cells(1, 2).Value = "Layout: " & GetLayoutModeName(mLayoutMode) _
                           & IIf(mLayoutMode = lmCurrentAtSlot, " (Slot=" & CStr(mAnchorSlot) & ")", "")
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(2, 2).Value = "Base Year: " & CStr(baseYear)
    
    If have Then
        Dim totalDays As Long, biz As Long
        totalDays = CLng(eRange - sRange) + 1
        biz = CountBusinessDays(sRange, eRange)
        If mLang = LangE Then
            ws.Cells(3, 2).Value = "Range: " & FmtDateOut(sRange) & " ~ " & FmtDateOut(eRange) _
                               & "  (" & CStr(biz) & " / " & CStr(totalDays) & ")"
        Else
            ws.Cells(3, 2).Value = "선택범위: " & FmtDateOut(sRange) & " ~ " & FmtDateOut(eRange) _
                                   & "  (" & CStr(biz) & " / " & CStr(totalDays) & ")"
        End If
        ws.Cells(3, 2).Font.Bold = True
    Else
        If mLang = LangE Then
            ws.Cells(3, 2).Value = "Range: (Nothing)"
        Else
            ws.Cells(3, 2).Value = "선택범위: (없음)"
        End If
        
    End If
    
    ' 폰트 기본
    With ws.Cells
        .Font.name = "Segoe UI"
        .Font.Size = 8
    End With
    
    ' === 12개월 렌더 ===
    Dim i As Long, y As Long, m As Long
    For i = 1 To 12
        y = Year(mapDate(i)): m = Month(mapDate(i))
        
        Dim rowBlock As Long, colBlock As Long
        rowBlock = (i - 1) \ GRID_COLS
        colBlock = (i - 1) Mod GRID_COLS
        
        Dim c0 As Long, r0 As Long
        c0 = COL0 + colBlock * (COLS_PER_MONTH + GAP_COLS)
        r0 = ROW0 + rowBlock * (ROWS_PER_MONTH + GAP_ROWS)
        
        ' --- 타이틀 ---
        With ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + COLS_PER_MONTH - 1))
            .Merge
            .Value = MonthTitleOut(y, m)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = IIf((y Mod 2) = 0, YearEvenBG, YearOddBG)
            .Font.Color = YearFG
            .RowHeight = 16
            With .Borders
                .LineStyle = xlContinuous
                .Color = CLR_GRID
            End With
        End With
        
         ' --- 요일 헤더 (주 시작/표기 스타일 반영) ---
        Dim j As Long, realDow As Long
        For j = 1 To 7
            With ws.Cells(r0 + ROWS_TITLE, c0 + (j - 1))
                .Value = WeekNameByPos(j)                 ' ← 요일 텍스트(Short/Full, K/E 반영)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = vbWhite
        
                realDow = DowByPos(j)                      ' 1=Sun .. 7=Sat(주 시작에 영향 없음)
                If realDow = vbSunday Then
                    .Font.Color = CLR_SUN_FG               ' 일요일 빨강
                ElseIf realDow = vbSaturday Then
                    .Font.Color = CLR_SAT_FG               ' 토요일 파랑
                Else
                    .Font.Color = CLR_TEXT                 ' 평일 기본색
                End If
        
                .RowHeight = 14
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = CLR_GRID
                End With
            End With
        Next
        
        ' --- 날짜 채우기 ---
        Dim firstD As Date: firstD = DateSerial(y, m, 1)
        Dim lastDay As Long: lastDay = Day(DateSerial(y, m + 1, 0))
        
        Dim startCol As Long: startCol = Weekday(firstD, FirstDOWParam())
        
         ' --- 날짜 채우기 (공휴일 메모 포함) ---
        Dim rr As Long, cc As Long, d As Long
        rr = 0: cc = startCol
        
        For d = 1 To lastDay
            Dim r As Long, c As Long
            r = r0 + ROWS_TITLE + ROWS_WEEK + rr
            c = c0 + (cc - 1)
        
            Dim dd As Date
            dd = DateSerial(y, m, d)
        
            With ws.Cells(r, c)
                Dim note As String
                note = ""
            
                .Value = d
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = CLR_BG
                .Font.Color = CLR_TEXT
                .Font.Bold = False
                .RowHeight = 14
        
                ' 주말 색
                If IsSundayDate(dd) Then
                    .Font.Color = CLR_SUN_FG
                    .Interior.Color = CLR_SUN_BG
                ElseIf IsSaturdayDate(dd) Then
                    .Font.Color = CLR_SAT_FG
                    .Interior.Color = CLR_SAT_BG
                End If
                
                If Not mTaskOverlay Then
                    ' 공휴일 메모 → note에 누적 (기존 .AddComment 삭제)
                    If IsHolidayWS(dd) Then
                        .Interior.Color = CLR_HOLI_BG
                        .Font.Color = CLR_SUN_FG
                        Dim holName As String
                        holName = GetHolidayNameIfAny(dd)
                        If Len(holName) > 0 Then
                            Dim holPrefix As String
                            holPrefix = IIf(mLang = LangE, "[Holiday] ", "[공휴일] ")
                            If Len(note) > 0 Then note = note & vbCrLf
                            note = note & holPrefix & holName
                        End If
                    End If
                End If
        
                ' 오늘 강조
                If dd = Date Then
                    .Interior.Color = CLR_TODAY_BG
                    .Font.Bold = True
                End If
        
                ' 선택 범위 오버레이
                If have Then
                    If dd >= sRange And dd <= eRange Then
                        If chkExcludeNonBiz.Value Then
                            If Not (IsWeekend(dd) Or IsHolidayWS(dd)) Then
                                .Interior.Color = CLR_SEL_BG
                                .Font.Color = CLR_SEL_FG
                                .Font.Bold = True
                            End If
                        Else
                            .Interior.Color = CLR_SEL_BG
                            .Font.Color = CLR_SEL_FG
                            .Font.Bold = True
                        End If
                    End If
                End If
                
                ' === ★ 오버레이 복제: UI DayBox의 모습(배경/글꼴/테두리/툴팁)을 그대로 시트에 반영 ===
                Dim tbUI As MSForms.TextBox
                Set tbUI = FindDayBoxByDate(dd)   ' UI 상 해당 날짜 칸
                If Not tbUI Is Nothing Then
                    On Error Resume Next
                    ' 1) 배경/글꼴(선택범위/오버레이가 이미 반영된 최종 색상)
                    .Interior.Color = tbUI.BackColor
                    .Font.Bold = tbUI.Font.Bold
                    .Font.Color = tbUI.ForeColor
                
                    ' 2) 오버레이 테두리도 반영(있을 때)
                    If tbUI.BorderStyle <> 0 Then
                        With .Borders
                            .LineStyle = xlContinuous
                            .Color = tbUI.BorderColor
                            .Weight = xlThin
                        End With
                    End If
                    
                    If mTaskOverlay Then
                        ' 3) 오버레이 툴팁을 메모에 누적
                        If Len(tbUI.ControlTipText) > 0 Then
                            If Len(tbUI.ControlTipText) = 10 Then
                            Else
                                If Len(note) > 0 Then note = note & vbCrLf
                                'note = note & IIf(mLang = LangE, "[Tasks] ", "[작업] ") & Mid(tbUI.ControlTipText, 11)
                                note = note & Mid(tbUI.ControlTipText, 11)
                            End If
                        End If
                    End If
                
'                    ' 3) 오버레이 툴팁을 메모에 누적
'                    If Len(tbUI.ControlTipText) > 0 Then
'                        If Len(note) > 0 Then note = note & vbCrLf
'                        note = note & IIf(mLang = LangE, "[Tasks] ", "[작업] ") & tbUI.ControlTipText
'                    End If
                    On Error GoTo 0
                End If
                
                ' === ★ 최종 메모 기록(공휴일+오버레이 합본) ===
                If Len(note) > 0 Then
                    On Error Resume Next
                    If Not .Comment Is Nothing Then .Comment.Delete
                    .AddComment note
                    .Comment.visible = False
                    With .Comment.Shape.TextFrame
                        .AutoSize = True
                        .Characters.Font.name = "맑은 고딕"
                        .Characters.Font.Size = 9
                    End With
                    On Error GoTo 0
                End If
        
                ' 테두리
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = CLR_GRID
                End With
            End With
        
            ' 다음 칸
            cc = cc + 1
            If cc > 7 Then cc = 1: rr = rr + 1: If rr >= ROWS_DAYS Then Exit For
        Next d
        
        ' 빈칸(달 외 영역)은 값/색을 비운 채 테두리만 얇게 유지(가독성)
        Dim rDays As Long, cDays As Long
        For rDays = 0 To ROWS_DAYS - 1
            For cDays = 1 To 7
                Dim rCell As Long: rCell = r0 + ROWS_TITLE + ROWS_WEEK + rDays
                Dim cCell As Long: cCell = c0 + cDays - 1
                With ws.Cells(rCell, cCell)
                    If Len(.Value) = 0 Then
                        .Interior.Color = CLR_BG
                        .Font.Color = CLR_TEXT
                        With .Borders
                            .LineStyle = xlContinuous
                            .Color = CLR_GRID
                        End With
                    End If
                End With
            Next cDays
        Next rDays
        
        ' 칼럼 폭 균일화(월 블록의 7열)
        For j = 0 To 6
            ws.Columns(c0 + j).ColumnWidth = 3.6
        Next
    Next i
    
    'ws.Columns(1).ColumnWidth = 2
    ws.Range("A:A,I:J,R:S,AA:AB").ColumnWidth = 2
    
    ws.Rows(1).EntireRow.AutoFit
    ws.Rows(2).EntireRow.AutoFit
    ws.Rows(3).EntireRow.AutoFit
    
    ActiveWindow.DisplayGridlines = False
    
    Application.ScreenUpdating = True
    ws.Activate
End Sub

Private Function GetLayoutModeName(ByVal m As eLayoutMode) As String
    Select Case m
        Case lmNormal:         GetLayoutModeName = "Normal (1~12)"
        Case lmCurrentFirst:   GetLayoutModeName = "Current First"
        Case lmCurrentLast:    GetLayoutModeName = "Current Last"
        Case lmCurrentAtSlot:  GetLayoutModeName = "Current @ Slot"
        Case Else:             GetLayoutModeName = "Unknown"
    End Select
End Function

' 전역 gHolidaySet(Dictionary: Key=Date, Val=Name)을 사용
Private Function IsHolidayWS(ByVal d As Date) As Boolean
    On Error Resume Next
    If gHolidaySet Is Nothing Then Exit Function
    IsHolidayWS = gHolidaySet.Exists(CDate(Int(CDbl(d))))
End Function

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Me.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Public Function RouteKey(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    Select Case KeyCode
        Case vbKeyPageDown
            If (Shift And SHIFT_MASK) <> 0 And (Shift And CTRL_MASK) <> 0 Then
                spnYear.Value = spnYear.Value + 15
            Else
                If (Shift And CTRL_MASK) <> 0 Then
                    spnYear.Value = spnYear.Value + 5
                Else
                    If (Shift And SHIFT_MASK) <> 0 Then
                        spnYear.Value = spnYear.Value + 10
                    Else
                        spnYear.Value = spnYear.Value + 1
                    End If
                End If
            End If
            RouteKey = True
        Case vbKeyPageUp
            If (Shift And SHIFT_MASK) <> 0 And (Shift And CTRL_MASK) <> 0 Then
                spnYear.Value = spnYear.Value - 15
            Else
                If (Shift And CTRL_MASK) <> 0 Then
                    spnYear.Value = spnYear.Value - 5
                Else
                    If (Shift And SHIFT_MASK) <> 0 Then
                        spnYear.Value = spnYear.Value - 10
                    Else
                        spnYear.Value = spnYear.Value - 1
                    End If
                End If
            End If
            RouteKey = True
        Case vbKeyEnd
            spnYear.Value = Year(Date)
            RouteKey = True
        Case vbKeyReturn
            btnApplyRange_Click
            RouteKey = True
        Case vbKeyEscape
            Unload Me
            RouteKey = True
        Case vbKeyDelete
            ClearSelectionUI
            RouteKey = True
        Case VK_OEM_PERIOD, vbKeyDecimal   ' "."
            SetToDate Date
            RouteKey = True
        Case VK_OEM_COMMA    ' ","
            SetFromDate Date
            RouteKey = True
    End Select
End Function
'
'' frmYearCalendar 내에 추가
'Public Sub OnMouseWheel(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)
'    ' dir: +1(위), -1(아래)로 들어옵니다.
'    If dir = 0 Then Exit Sub
'
'    Dim stepVal As Long: stepVal = 1
'    ' 기존 RouteKey(PageUp/Down) 규칙과 동일한 보간:
'    '  Shift+Ctrl: 15,  Ctrl: 5,  Shift: 10,  기본: 1
'    If isShift And isCtrl Then
'        stepVal = 15
'    ElseIf isCtrl Then
'        stepVal = 5
'    ElseIf isShift Then
'        stepVal = 10
'    Else
'        stepVal = 1
'    End If
'
'    On Error Resume Next
'    If dir > 0 Then
'        spnYear.Value = spnYear.Value - stepVal
'    Else
'        spnYear.Value = spnYear.Value + stepVal
'    End If
'    ' spnYear_Change 에서 이미 RenderAllMonths/PaintRange 호출하므로 추가 호출 불필요
'End Sub

Private Sub SetFromDate(d As Date)
    mFrom = d
    mHasFrom = True

    ' From/To 정규화(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' 상단 표시
    txtFromSel.text = IIf(mHasFrom, Format$(mFrom, "yyyy-mm-dd"), "")

    PaintRange        ' 선택 구간 다시 칠하기
    UpdateRangeInfo   ' "Business Day / 총일수" 갱신
    ApplySelectedRangeOverlay
End Sub

Public Sub SetToDate(d As Date)
    mTo = d
    mHasTo = True

    ' From/To 정규화(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' 상단 표시
    txtToSel.text = IIf(mHasTo, Format$(mTo, "yyyy-mm-dd"), "")

    PaintRange        ' 선택 구간 다시 칠하기
    UpdateRangeInfo   ' "Business Day / 총일수" 갱신
    ApplySelectedRangeOverlay
End Sub


' frmYearCalendar 내 임의 위치(모듈 범위) 추가
Public Sub HandleMonthLabelMouse(ByVal L As MSForms.Label, ByVal Button As Integer)
    Dim idx As Long
    Dim baseYear As Long, curM As Long
    Dim mapDate() As Date
    Dim y As Long, m As Long
    Dim d As Date

    ' 레이블 이름: "lblMonth" & 인덱스(1..12)
    idx = CLng(Val(Mid$(L.name, 9)))
    If idx < 1 Or idx > 12 Then Exit Sub

    ' 현재 화면의 12칸이 가리키는 실제 (연/월) 매핑
    baseYear = CLng(Val(txtBaseYear.text))
    curM = Month(Date)
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot

    y = Year(mapDate(idx))
    m = Month(mapDate(idx))

    Select Case Button
        Case 1 ' Left → From = 그 달의 1일
            d = DateSerial(y, m, 1)
            mFrom = d: mHasFrom = True
        Case 2 ' Right → To = 그 달의 말일
            d = DateSerial(y, m + 1, 0)
            mTo = d: mHasTo = True
        Case Else
            Exit Sub
    End Select

    ' From ≤ To 보정
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim t As Date: t = mFrom: mFrom = mTo: mTo = t
        End If
    End If

    ' UI 반영 및 칠하기
    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")

    PaintRange
    UpdateRangeInfo
    ApplySelectedRangeOverlay
End Sub

' frmYearCalendar
Public Sub HandleMonthLabelDouble(ByVal L As MSForms.Label)
    Dim idx As Long, baseYear As Long, curM As Long, mapDate() As Date
    Dim y As Long, m As Long
    idx = CLng(Val(Mid$(L.name, 9)))
    If idx < 1 Or idx > 12 Then Exit Sub
    baseYear = CLng(Val(txtBaseYear.text))
    curM = Month(Date)
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot
    y = Year(mapDate(idx)): m = Month(mapDate(idx))
    mFrom = DateSerial(y, m, 1): mHasFrom = True
    mTo = DateSerial(y, m + 1, 0): mHasTo = True
    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")
    
    PaintRange
    UpdateRangeInfo
    ApplySelectedRangeOverlay
End Sub

'-- 2025.09.14 17:49 Add
Private Sub LoadI18NAndFormats()
    Dim s As String
    s = GetSetting(REG_APP, REG_SEC_I18N, REG_KEY_LANG, "K")
    mLang = IIf(UCase$(s) = "E", LangE, LangK)

    ' 날짜 포맷(공용; 필요시 분리 가능)
    mFmtDate = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_DATE, _
                 IIf(mLang = LangE, DEF_FMT_DATE_E, DEF_FMT_DATE_K))

    ' ===== 월 타이틀 포맷(언어별 분리 저장) =====
    ' 언어별 타이틀 포맷 로드
    Dim rawK As String, rawE As String, legacy As String
    rawK = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE_K, DEF_FMT_TITLE_K)
    rawE = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE_E, DEF_FMT_TITLE_E)
    
    ' (옵션) 레거시 키가 있으면 초기 이관
    legacy = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE, "")
    If legacy <> "" Then
        If rawK = DEF_FMT_TITLE_K Then rawK = legacy
        If rawE = DEF_FMT_TITLE_E Then rawE = legacy
    End If
    
    ' 부적합 서식 방어
    mFmtMonthTitleK = SanitizeMonthTitleFmt(rawK, False)
    mFmtMonthTitleE = SanitizeMonthTitleFmt(rawE, True)
    
    ' 3) 현재 언어에 맞게 선택
    mFmtMonthTitle = IIf(mLang = LangE, mFmtMonthTitleE, mFmtMonthTitleK)

    ' --- WeekStart / WeekNameStyle (이전 답변에서 추가한 부분 유지) ---
    Dim ws As String, wn As String
    ws = GetSetting(REG_APP, REG_SEC_I18N, REG_KEY_WEEK_START, "Sun")
    wn = GetSetting(REG_APP, REG_SEC_I18N, REG_KEY_WEEKNAME_STYLE, "Short")
    mWeekStart = IIf(UCase$(ws) = "MON", WeekMon, WeekSun)
    mWeekNameStyle = IIf(UCase$(wn) = "FULL", WkFull, WkShort)
End Sub

Private Function SanitizeMonthTitleFmt(ByVal s As String, ByVal isEnglish As Boolean) As String
    Dim t As String: t = Trim$(s)
    If t = "" Then GoTo def
    ' 일자 토큰(d/D)이 들어간 경우는 월 타이틀용으로 부적합 → 기본값으로
    If InStr(1, LCase$(t), "d", vbTextCompare) > 0 Then GoTo def
    SanitizeMonthTitleFmt = t
    Exit Function
def:
    SanitizeMonthTitleFmt = IIf(isEnglish, DEF_FMT_TITLE_E, DEF_FMT_TITLE_K)
End Function

Private Sub SaveLangToRegistry()
    SaveSetting REG_APP, REG_SEC_I18N, REG_KEY_LANG, IIf(mLang = LangE, "E", "K")
End Sub

' 언어별 요일 약칭
Private Function WeekNameByIndex(ByVal idx As Long) As String
    ' idx: 1=Sun ~ 7=Sat
    If mLang = LangE Then
        WeekNameByIndex = Split("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")(idx - 1)
    Else
        WeekNameByIndex = Split("일,월,화,수,목,금,토", ",")(idx - 1)
    End If
End Function

' 월 타이틀 문자열
Private Function MonthTitleOut(ByVal y As Long, ByVal m As Long) As String
    MonthTitleOut = Format$(DateSerial(y, m, 1), mFmtMonthTitle)
End Function

' 날짜 출력 문자열 (텍스트박스/반환용)
Private Function FmtDateOut(ByVal d As Date) As String
    FmtDateOut = Format$(d, mFmtDate)
End Function

' Excel 셀에 날짜 적용(값+표시형식) ? 모든 셀 기록 지점에서 이걸 호출
Private Sub ApplyDateToCell(ByVal tgt As Range, ByVal d As Date)
    If tgt Is Nothing Then Exit Sub
    tgt.Value = d
    On Error Resume Next
    tgt.NumberFormatLocal = mFmtDate   ' 대부분 Excel과 동일 토큰. 지역 포맷 사용
    On Error GoTo 0
End Sub

' UI 정적 텍스트(버튼/체크박스 캡션 등) 한/영 반영
Private Sub ApplyStaticUIStrings()
    On Error Resume Next

    ' === 우상단 K/E 토글 체크박스 ===
    chkKE.Caption = "K/E"

    ' 버튼/체크박스/옵션 등 (필요한 범위만 예시)
    If mLang = LangE Then
        btnCurrYear.Caption = "ThisYear"
        btnApplyRange.Caption = "Apply"
        btnClear.Caption = "Clear"
        btnClose.Caption = "Close"
        btnExport.Caption = "Export"
        btnConfig.Caption = "Set"

        optNormal.Caption = "1→12"
        optCurrentFirst.Caption = "Current First"
        optCurrentLast.Caption = "Current Last"
        optCurrentAtSlot.Caption = "Current @Slot"

        chkFromOnly.Caption = "From Only"
        chkShowKeep.Caption = "Keep form open"
        chkExcludeNonBiz.Caption = "Exclude Sat/Sun/Holidays in range color"
        chkAutoNextRow.Caption = "Go↓"

        lblFromTo.Caption = "From~To"
        lblBaseYear.Caption = "Year"
        lblSlot.Caption = "Slot"
        
        btnTaskLoad.Caption = "Load"
        btnTaskSaveAppend.Caption = "Save Append"
        
        btnTaskGet.Caption = "Import from selected range"
        btnTaskExport.Caption = "Export To Sheet"
        chkTaskOverlay.Caption = "Overlay tasks on Calendar"
        chkTaskLinkSel.Caption = "Set From~To from Selected Task"
        chkTaskShowAll.Caption = "Show All" & vbLf & "(If off, only current calendar-range tasks)"
        btnRemoveTaskYear.Caption = "Delete All (Current Year)"
        btnTaskDeleteSelected.Caption = "Delete Selected Task"
        
        btnTaskAdd.Caption = "Add"
        btnTaskUpdate.Caption = "Update"
        btnTaskDelete.Caption = "Delete"
        btnTaskSort.Caption = "Sort"
        

    Else
        btnCurrYear.Caption = "올해"
        btnApplyRange.Caption = "적용(반환)"
        btnClear.Caption = "Clear"
        btnClose.Caption = "닫기"
        btnExport.Caption = "출력"
        btnConfig.Caption = "설정"

        optNormal.Caption = "1월→12월"
        optCurrentFirst.Caption = "현재월을 시작에"
        optCurrentLast.Caption = "현재월을 끝에"
        optCurrentAtSlot.Caption = "현재월 @Slot"

        chkFromOnly.Caption = "From Only"
        chkShowKeep.Caption = "창 유지"
        chkExcludeNonBiz.Caption = "범위표시에 휴일 제외"
        chkAutoNextRow.Caption = "아래이동"

        lblFromTo.Caption = "From~To"
        lblBaseYear.Caption = "연도"
        lblSlot.Caption = "슬롯"
        
        btnTaskLoad.Caption = "불러오기"
        btnTaskSaveAppend.Caption = "추가 저장"
        
        btnTaskGet.Caption = "시트 선택영역에서 불러오기"
        btnTaskExport.Caption = "시트에 출력"
        chkTaskOverlay.Caption = "Task 일정을 Calendar에 표시"
        chkTaskLinkSel.Caption = "Task 선택시 From~To에 반영"
        chkTaskShowAll.Caption = "전체 보기" & vbLf & "(해제시 現 Calendar 구간 Task만 표시)"
        btnRemoveTaskYear.Caption = "현재 연도 모두 삭제"
        btnTaskDeleteSelected.Caption = "선택 항목 삭제"
        
        btnTaskAdd.Caption = "추가"
        btnTaskUpdate.Caption = "수정"
        btnTaskDelete.Caption = "삭제"
        btnTaskSort.Caption = "정렬"
        
    End If
End Sub

' 요일 헤더/타이틀 등 재적용
Private Sub ReapplyLanguageOnCalendar()
    Dim i As Long, j As Long
    ' 월 타이틀
    For i = 1 To 12
        Dim y As Long, m As Long, baseYear As Long, curM As Long, mapDate() As Date
        baseYear = CLng(Val(txtBaseYear.text))
        curM = Month(Date)
        BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot
        y = Year(mapDate(i)): m = Month(mapDate(i))
        If Not lblMonth(i) Is Nothing Then lblMonth(i).Caption = MonthTitleOut(y, m)
        ' 요일 헤더
        For j = 1 To 7
            If Not lblWeek(i, j) Is Nothing Then
                lblWeek(i, j).Caption = WeekNameByPos(j)
                Dim dow As Long: dow = DowByPos(j)
                If dow = vbSunday Then
                    lblWeek(i, j).ForeColor = mClrSunFg
                ElseIf dow = vbSaturday Then
                    lblWeek(i, j).ForeColor = mClrSatFg
                Else
                    lblWeek(i, j).ForeColor = mClrWeekFg
                End If
            End If
        Next
    Next
End Sub

'----------------------------------------
Private Sub chkKE_Click()
    mLang = IIf(chkKE.Value, LangE, LangK)   ' 체크=영문, 해제=한글
    SaveLangToRegistry
    ' 언어 바뀌면 포맷 기본값도 한 번 동기화(사용자가 설정해둔 값이 있으면 그대로 유지)
    If mFmtDate = "" Then mFmtDate = IIf(mLang = LangE, DEF_FMT_DATE_E, DEF_FMT_DATE_K)
    If mFmtMonthTitle = "" Then mFmtMonthTitle = IIf(mLang = LangE, DEF_FMT_TITLE_E, DEF_FMT_TITLE_K)

    ApplyStaticUIStrings
    ' 월 타이틀/요일 재적용 + 셀 도색 유지
    ReapplyLanguageOnCalendar
    PaintRange
    
    SaveLangToRegistry
    LoadI18NAndFormats        ' ← 언어 바꾸면 월 타이틀 포맷도 언어별 값으로 재로딩
    ReapplyLanguageOnCalendar ' ← 월 타이틀/요일 캡션 다시 반영
    
End Sub

Private Sub btnConfig_Click()
    frmDateFormatConfig.Show vbModal
    ' 사용자가 저장했다면 레지스트리에서 다시 읽고 반영
    LoadI18NAndFormats
    ApplyStaticUIStrings
    RenderAllMonths          ' ← 주 시작 바뀌면 배치가 달라지므로 재렌더
    ReapplyLanguageOnCalendar
    ' 선택 텍스트도 새 서식으로 다시 표시
    If mHasFrom Then txtFromSel.text = FmtDateOut(mFrom)
    If mHasTo Then txtToSel.text = FmtDateOut(mTo)
    PaintRange
End Sub

' 주 시작 요일을 VbDayOfWeek로 환산
Private Function FirstDOWParam() As VbDayOfWeek
    FirstDOWParam = IIf(mWeekStart = WeekMon, vbMonday, vbSunday)
End Function

' 화면상의 위치(1..7) → 실제 요일번호(1=Sun .. 7=Sat)
Private Function DowByPos(ByVal pos As Long) As Long
    ' WeekSun: pos=1→Sun,2→Mon,...7→Sat
    ' WeekMon: pos=1→Mon(2), ..., 6→Sat(7), 7→Sun(1)
    If mWeekStart = WeekSun Then
        DowByPos = pos
    Else
        DowByPos = ((pos Mod 7) + 1)   ' 1→2, 2→3, ..., 6→7, 7→1
    End If
End Function

' 요일명(언어/스타일 반영)
Private Function WeekNameByDow(ByVal dow As Long) As String
    If mLang = LangE Then
        If mWeekNameStyle = WkFull Then
            WeekNameByDow = Split("Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", ",")(dow - 1)
        Else
            WeekNameByDow = Split("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")(dow - 1)
        End If
    Else
        ' 한글은 기존 약칭 유지(원하시면 '일요일,월요일,...'로 Full 추가 가능)
        If mWeekNameStyle = WkFull Then
            WeekNameByDow = Split("일요일,월요일,화요일,수요일,목요일,금요일,토요일", ",")(dow - 1)
        Else
            WeekNameByDow = Split("일,월,화,수,목,금,토", ",")(dow - 1)
        End If
    End If
End Function

' 화면 위치(1..7) → 요일명
Private Function WeekNameByPos(ByVal pos As Long) As String
    WeekNameByPos = WeekNameByDow(DowByPos(pos))
End Function

' 날짜 기반 주말 판정(주 시작과 무관)
Private Function IsSundayDate(ByVal d As Date) As Boolean
    IsSundayDate = (Weekday(d, vbSunday) = vbSunday)
End Function
Private Function IsSaturdayDate(ByVal d As Date) As Boolean
    IsSaturdayDate = (Weekday(d, vbSunday) = vbSaturday)
End Function

Private Sub btnHelp_Click()
    ' 모델리스로 두고 달력과 동시에 참고하려면 vbModeless
    frmYearCalHelp.Show vbModeless
End Sub

' 연도 설정을 한 곳에서만 처리: mYear, txtBaseYear, spnYear.Value 동기화 + (옵션)렌더
Private Sub SetBaseYear(ByVal y As Long, Optional ByVal doRender As Boolean = True)
    Dim newY As Long
    newY = y
    If newY < 1901 Then newY = 1901
    If newY > 9998 Then newY = 9998

    If mSyncingYear Then Exit Sub
    mSyncingYear = True
    On Error Resume Next

    ' 스핀 범위 보장
    If spnYear.Min <> 1901 Then spnYear.Min = 1901
    If spnYear.Max <> 9998 Then spnYear.Max = 9998

    ' 실제 변경 여부
    Dim changed As Boolean
    changed = (mYear <> newY)

    mYear = newY
    If NzCStr(txtBaseYear.text) <> CStr(newY) Then txtBaseYear.text = CStr(newY)
    If spnYear.Value <> newY Then spnYear.Value = newY

    mSyncingYear = False
    On Error GoTo 0

    If doRender And changed Then
        RenderAllMonths
        PaintRange
    End If
End Sub

'===== 2025.09.20 Task Panel =====
' === 초기화 시점에 호출 (기존 Initialize 끝부분에 이어서 넣어주세요) ===
Private Sub InitTaskPanel()
    ' "NOHOOK" 태그가 Panel 컨트롤에 지정되어 있어야 함
    ' ListBox 컬럼 설정
    With Me.lstTask
        .ColumnCount = 3
        .ColumnHeads = False
        ' 3열: Name | From | To
        '.ColumnWidths = "100;20;20"
        .BoundColumn = 0
        .MultiSelect = fmMultiSelectExtended   ' ★ 다중 선택 허용
    End With

    ' 폼 너비/토글폭 설정 (디자인 기준으로 보정하세요)
    mBaseWidth = 700 ' 접힘 상태의 기본폭 (필요시 폼 현재 Width로 대체)
    mTaskWidth = 975 ' 펼친 상태의 폭
    
    ' === ★ 레지스트리에서 최종 상태 로드 ===
    mTaskVisible = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_VISIBLE, True)
    mTaskOverlay = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_OVERLAY, False)
    mTaskLinkSel = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_LINKSEL, True)
    
    EnsureTaskPanelVisible mTaskVisible

    ' 체크박스 상태 초기화(레지스트리 등에서 불러와도 됨)
    Me.chkTaskOverlay.Value = mTaskOverlay
    Me.chkTaskLinkSel.Value = mTaskLinkSel

    ' 캘린더 DayBox 매핑 준비
    BuildDayBoxMapFromGrid
    
    ' ▼ 마지막에 기준년도 Task 자동 로드 & 오버레이 반영
    LoadTasksForVisibleRangeAndOverlay
    
    ' ... 기존 코드 ...
    ' === 전체 보기 초기 상태(레지스트리 복원) ===
    Me.chkTaskShowAll.Value = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_SHOWALL, False)
    
    ' === 처음 로드 ===
    LoadCategoryList
    RefreshTaskListAndOverlay   ' (아래 함수)
    
End Sub

Private Sub EnsureTaskPanelVisible(ByVal visible As Boolean)
    mTaskVisible = visible
    Me.Width = IIf(visible, mTaskWidth, mBaseWidth)
    Dim en As Boolean: en = visible
    ' 패널 컨트롤 포커스/탭/활성 제한
    PanelSetEnabled en
End Sub

Private Sub PanelSetEnabled(ByVal en As Boolean)
    Dim ctl As MSForms.control
    For Each ctl In Me.Controls
        If IsTaskPanelControl(ctl) Then
            On Error Resume Next
            ctl.Enabled = en
            ' 탭 정지 방지
            If HasProp(ctl, "TabStop") Then
                CallByName ctl, "TabStop", VbLet, en
            End If
            On Error GoTo 0
        End If
    Next
End Sub

Private Function IsTaskPanelControl(ByVal ctl As MSForms.control) As Boolean
    ' 디자인시 Task Panel 영역 컨트롤 전부 Tag="NOHOOK" 부여
    IsTaskPanelControl = (InStr(1, GetCtlTag(ctl), "NOHOOK", vbTextCompare) > 0)
End Function

Private Function GetCtlTag(ByVal ctl As Object) As String
    On Error Resume Next
    GetCtlTag = Trim$(ctl.Tag)
End Function

Private Function HasProp(obj As Object, propName As String) As Boolean
    On Error GoTo EH
    Dim tmp: tmp = CallByName(obj, propName, VbGet)
    HasProp = True
    Exit Function
EH:
    HasProp = False
End Function

'=== tbDay(12x6x7) 순회하여 "yyyy-mm-dd" → TextBox 매핑 ===
Private Sub BuildDayBoxMapFromGrid()
    Dim m As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox, key As String

    Set DayBoxByDate = CreateObject("Scripting.Dictionary")
    If OrigToolTip Is Nothing Then Set OrigToolTip = CreateObject("Scripting.Dictionary")

    For m = LBound(tbDay, 1) To UBound(tbDay, 1)
        For r = LBound(tbDay, 2) To UBound(tbDay, 2)
            For c = LBound(tbDay, 3) To UBound(tbDay, 3)
                Set tb = tbDay(m, r, c)
                If Not tb Is Nothing Then
                    'If TryParseYMD(GetCtlTag(tb), key) Then
                    If TryGetDateKeyFromTag(GetCtlTag(tb), key) Then
                        If Not DayBoxByDate.Exists(key) Then
                            DayBoxByDate.add key, tb
                        Else
                            Set DayBoxByDate(key) = tb
                        End If
                        ' 툴팁 원본만 캐시
                        If OrigToolTip.Exists(key) Then
                            OrigToolTip(key) = tb.ControlTipText
                        Else
                            OrigToolTip.add key, tb.ControlTipText
                        End If
                        ' 기본은 무테 규칙: 여기서는 BorderStyle 관여 X
                    End If
                End If
            Next c
        Next r
    Next m
End Sub

'=== Tag가 "yyyy-mm-dd" 이면 True, 표준키 반환 ===
Private Function TryGetDateKeyFromTag(ByVal s As String, ByRef key As String) As Boolean
    On Error GoTo EH
    TryGetDateKeyFromTag = False: key = ""
    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        Dim d As Date
        If TryParseYMD(s, d) Then
            key = Format$(d, "yyyy-mm-dd")
            TryGetDateKeyFromTag = True
        End If
    End If
    Exit Function
EH:
    TryGetDateKeyFromTag = False: key = ""
End Function

' Name이 "tbD_" 로 시작하면 뒤에서 날짜를 뽑아 "yyyy-mm-dd"로 반환
' 허용 예: tbD_20250107, tbD_2025-01-07, tbD_2025_1_7 등
Private Function TryGetDateKeyFromName(ByVal nm As String, ByRef key As String) As Boolean
    On Error GoTo EH
    TryGetDateKeyFromName = False: key = ""
    If Left$(nm, 4) <> "tbD_" Then Exit Function

    Dim rest As String: rest = Mid$(nm, 5)
    ' 1) 8자리 숫자(yyyymmdd)
    If Len(rest) = 8 And IsNumeric(rest) Then
        key = Format$(DateSerial(CLng(Left$(rest, 4)), _
                                 CLng(Mid$(rest, 5, 2)), _
                                 CLng(Mid$(rest, 7, 2))), "yyyy-mm-dd")
        TryGetDateKeyFromName = True
        Exit Function
    End If
    ' 2) 구분자 섞임: 숫자만 추출 후 3토큰(y,m,d)
    Dim i As Long, ch As String, buf As String
    For i = 1 To Len(rest)
        ch = Mid$(rest, i, 1)
        buf = buf & IIf(ch Like "[0-9]", ch, " ")
    Next
    Dim tokens() As String
    tokens = Split(Application.WorksheetFunction.Trim(buf), " ")
    If UBound(tokens) >= 2 Then
        key = Format$(DateSerial(CLng(tokens(0)), CLng(tokens(1)), CLng(tokens(2))), "yyyy-mm-dd")
        TryGetDateKeyFromName = True
    End If
    Exit Function
EH:
    TryGetDateKeyFromName = False: key = ""
End Function

Private Function FindDayBox(ByVal d As Date) As MSForms.TextBox
    Dim key As String: key = Format$(d, "yyyy-mm-dd")
    If Not DayBoxByDate Is Nothing Then
        If DayBoxByDate.Exists(key) Then
            Set FindDayBox = DayBoxByDate(key)
        End If
    End If
End Function

'=== 토글 버튼 ===
Private Sub btnTaskToggle_Click()
    Dim newVisible As Boolean
    newVisible = Not mTaskVisible
    EnsureTaskPanelVisible newVisible
    btnTaskToggle.Caption = IIf(mTaskVisible, "◀", "▶") ' 펼쳐진 상태면 닫는 아이콘(◀), 접힌 상태면 여는 아이콘(▶)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_VISIBLE, mTaskVisible
End Sub


Private Function GetCurrentYearForTask() As Long
    ' 기존 폼의 "기준연도" 보유 로직에 맞게 값을 반환하세요.
    ' 예시: txtYear 라는 TextBox가 있다고 가정
    On Error Resume Next
    Dim y As Long
    y = CLng(Me.txtBaseYear.Value)
    If y < 1900 Or y > 9999 Then y = Year(Date)
    GetCurrentYearForTask = y
End Function

' === Task Panel: [추가 저장] ===
Private Sub btnTaskSaveAppend_Click()
    On Error GoTo EH

    Dim cat As String: cat = SelectedCategoryName
    Dim cur As Collection: Set cur = LoadAllTasks_File_Cat(cat)         ' 기존 전체
    Dim add As Collection: Set add = TasksFromListBox(Me.lstTask)       ' 현재 목록

    ' 1) 중복 제거(이름/From/To 동일시 중복)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long, t As clsTaskItem, k As String
    ' 기존 → dict
    For i = 1 To cur.Count
        Set t = cur(i)
        k = TripleKey(t)
        If Not dict.Exists(k) Then dict.add k, t
    Next
    ' 추가 → dict (없을 때만)
    For i = 1 To add.Count
        Set t = add(i)
        k = TripleKey(t)
        If Not dict.Exists(k) Then dict.add k, t
    Next

    ' 2) dict → Collection
    Dim merged As New Collection, v
    For Each v In dict.items
        merged.add v
    Next

    ' 3) 정렬(From → To → Name)
    Set merged = SortTasksByFromToName(merged)

    ' 4) 저장(카테고리 파일 + 로그)
    SaveAllTasks_File_Cat cat, merged

    RefreshTaskListAndOverlay
    MsgBox "추가 저장 완료.", vbInformation
    Exit Sub
EH:
    MsgBox "추가 저장 중 오류: " & Err.Description, vbExclamation
End Sub


' (이름|From|To) 키 ? 이름 대소문자 무시, 날짜는 yyyy-mm-dd 고정
Private Function TaskKey3(ByVal t As clsTaskItem) As String
    Dim nm As String: nm = LCase$(Trim$(NzCStr(t.TaskName)))
    Dim f As String:  f = Format$(t.FromDate, "yyyy-mm-dd")
    Dim z As String
    If t.HasTo Then
        z = Format$(t.ToDate, "yyyy-mm-dd")
    Else
        z = ""                             ' To 없음은 빈 문자열로 비교
    End If
    TaskKey3 = nm & "|" & f & "|" & z
End Function

' To가 없으면 정렬 비교용 To는 From으로 사용
Private Function ToForSort(ByVal t As clsTaskItem) As Date
    If t.HasTo Then ToForSort = t.ToDate Else ToForSort = t.FromDate
End Function

' a<b:-1, a>b:1, 같음:0 (From → To → Name)
Private Function TaskCmp(a As clsTaskItem, b As clsTaskItem) As Long
    If a.FromDate < b.FromDate Then TaskCmp = -1: Exit Function
    If a.FromDate > b.FromDate Then TaskCmp = 1:  Exit Function

    Dim at As Date, bt As Date
    at = ToForSort(a): bt = ToForSort(b)
    If at < bt Then TaskCmp = -1: Exit Function
    If at > bt Then TaskCmp = 1:  Exit Function

    Dim na As String, nb As String
    na = LCase$(Trim$(NzCStr(a.TaskName)))
    nb = LCase$(Trim$(NzCStr(b.TaskName)))
    If na < nb Then
        TaskCmp = -1
    Else
        If na > nb Then
            TaskCmp = 1
        Else
            TaskCmp = 0
        End If
    End If
End Function

' 안전 정렬: 0/1개 예외 처리 + 삽입정렬
' 정렬: From → To → Name
Private Function SortTasksByFromToName(ByVal tasks As Collection) As Collection
    Dim out As New Collection
    Dim n As Long, i As Long, j As Long

    If tasks Is Nothing Then Set SortTasksByFromToName = out: Exit Function
    n = tasks.Count
    If n <= 1 Then
        For i = 1 To n: out.add tasks(i): Next
        Set SortTasksByFromToName = out
        Exit Function
    End If

    Dim arr() As clsTaskItem
    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = tasks(i)
    Next

    ' 삽입 정렬 (j 경계 먼저 확인 → 그 다음 비교)
    For i = 2 To n
        Dim cur As clsTaskItem
        Set cur = arr(i)
        j = i - 1

        Do While j >= 1
            ' j가 1 미만이면 비교하지 않음(단락평가 대체)
            If TaskCmp(arr(j), cur) > 0 Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = cur
    Next

    For i = 1 To n
        out.add arr(i)
    Next
    Set SortTasksByFromToName = out
End Function


'=== 추가 ===
Private Sub btnTaskAdd_Click()
    Dim sName As String: sName = Trim$(Me.txtTaskName.text)
    Dim sFrom As String: sFrom = Trim$(Me.txtTaskFrom.text)
    Dim sTo   As String: sTo = Trim$(Me.txtTaskTo.text)

    Dim dF As Date, dT As Date
    Dim okF As Boolean, okT As Boolean

    okF = TryParseDate(sFrom, dF)
    okT = (Len(sTo) > 0 And TryParseDate(sTo, dT))

    If Not okF Then
        MsgBox "시작일(From)이 올바르지 않습니다. (yyyy-mm-dd)", vbExclamation
        Exit Sub
    End If
    If okT And dT < dF Then
        MsgBox "종료일(To)은 시작일보다 빠를 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 비교는 정규화된 문자열로 수행
    Dim fromYMD As String, toYMD As String
    fromYMD = Format$(dF, "yyyy-mm-dd")
    toYMD = IIf(okT, Format$(dT, "yyyy-mm-dd"), "")

    ' 중복 검사
    Dim dupIdx As Long
    dupIdx = FindTaskRow(Me.lstTask, sName, fromYMD, toYMD)

    If dupIdx >= 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "동일한 항목이 이미 존재합니다." & vbCrLf & _
            "Task: " & IIf(Len(sName) = 0, "(이름 없음)", sName) & vbCrLf & _
            "From: " & fromYMD & vbCrLf & _
            "To  : " & IIf(Len(toYMD) = 0, "(없음)", toYMD) & vbCrLf & vbCrLf & _
            "그래도 추가하시겠습니까?", _
            vbExclamation + vbYesNo, "중복 확인")
        If resp = vbNo Then Exit Sub
    End If

    ' 추가
    Me.lstTask.AddItem
    Me.lstTask.list(Me.lstTask.ListCount - 1, 0) = sName
    Me.lstTask.list(Me.lstTask.ListCount - 1, 1) = fromYMD
    Me.lstTask.list(Me.lstTask.ListCount - 1, 2) = toYMD

    ' (옵션) 오버레이 켜져 있으면 즉시 갱신
    ' RefreshTaskOverlayIfOn
End Sub

'=== 변경(업데이트) ===
Private Sub btnTaskUpdate_Click()
    Dim r As Long: r = Me.lstTask.ListIndex
    If r < 0 Then
        MsgBox "변경할 항목을 선택하세요.", vbExclamation
        Exit Sub
    End If

    Dim sName As String: sName = Trim$(Me.txtTaskName.text)
    Dim sFrom As String: sFrom = Trim$(Me.txtTaskFrom.text)
    Dim sTo As String:   sTo = Trim$(Me.txtTaskTo.text)

    Dim dF As Date, dT As Date, okF As Boolean, okT As Boolean
    okF = TryParseDate(sFrom, dF)
    okT = TryParseDate(sTo, dT)

    If Not okF Then
        MsgBox "시작일(From)이 올바르지 않습니다. (yyyy-mm-dd)", vbExclamation
        Exit Sub
    End If
    If okT And dT < dF Then
        MsgBox "종료일(To)은 시작일보다 빠를 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Me.lstTask.list(r, 0) = sName
    Me.lstTask.list(r, 1) = Format$(dF, "yyyy-mm-dd")
    Me.lstTask.list(r, 2) = IIf(okT, Format$(dT, "yyyy-mm-dd"), "")
End Sub

'=== 삭제 ===
Private Sub btnTaskDelete_Click()
    Dim r As Long: r = Me.lstTask.ListIndex
    If r < 0 Then
        MsgBox "삭제할 항목을 선택하세요.", vbExclamation
        Exit Sub
    End If
    Me.lstTask.RemoveItem r
End Sub

'=== 정렬 (From 기준) ===
Private Sub btnTaskSort_Click()
    Dim tasks As Collection
    Set tasks = TasksFromListBox(Me.lstTask)
    Set tasks = SortTasksByFromDate(tasks)
    FillListBoxFromTasksSafe Me.lstTask, tasks
End Sub

'=== 시트에서 불러오기 (선택영역 3열 x n행) ===
Private Sub btnTaskGet_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "3열 x n행의 범위를 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    If rng.Columns.Count < 3 Then
        MsgBox "선택 영역은 최소 3개 열이어야 합니다. (Name, From, To)", vbExclamation
        Exit Sub
    End If

    Dim r As Range
    For Each r In rng.Rows
        Dim sName As String, sFrom As String, sTo As String
        sName = Trim$(NzCStr(r.Cells(1, 1).Value))
        sFrom = Trim$(NzCStr(r.Cells(1, 2).Value))
        sTo = Trim$(NzCStr(r.Cells(1, 3).Value))

        If Len(sFrom) > 0 Then
            ' 유효성은 추가 버튼과 동일 규칙 사용
            Dim dF As Date, dT As Date, okF As Boolean, okT As Boolean
            okF = TryParseDate(sFrom, dF)
            okT = TryParseDate(sTo, dT)
            If okF Then
                Me.lstTask.AddItem
                Me.lstTask.list(Me.lstTask.ListCount - 1, 0) = sName
                Me.lstTask.list(Me.lstTask.ListCount - 1, 1) = Format$(dF, "yyyy-mm-dd")
                Me.lstTask.list(Me.lstTask.ListCount - 1, 2) = IIf(okT, Format$(dT, "yyyy-mm-dd"), "")
            End If
        End If
    Next
End Sub

'=== 시트에 출력 (New Sheet) ===
Private Sub btnTaskExport_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = "Tasks_" & Format$(Now, "yyyymmdd_hhnnss")

    ws.Range("A1").Value = "Task Name"
    ws.Range("B1").Value = "From"
    ws.Range("C1").Value = "To"

    Dim r As Long, row As Long: row = 2
    For r = 0 To Me.lstTask.ListCount - 1
        ws.Cells(row, 1).Value = NzCStr(Me.lstTask.list(r, 0))
        ws.Cells(row, 2).Value = NzCStr(Me.lstTask.list(r, 1))
        ws.Cells(row, 3).Value = NzCStr(Me.lstTask.list(r, 2))
        row = row + 1
    Next

    ws.Columns("A:C").AutoFit
    MsgBox "새 시트에 출력했습니다: " & ws.name, vbInformation
End Sub

'=== 오버레이 체크박스 ===
Private Sub chkTaskOverlay_Click()
    mTaskOverlay = (Me.chkTaskOverlay.Value = True)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_OVERLAY, mTaskOverlay
    
    ReapplyOverlayAndSelection (False)

'    ' 최신 그리드 매핑
'    BuildDayBoxMapFromGrid
'
'    ' 1) 항상 먼저 완전 초기화
'    ClearAllDayBoxOverlay
'
'    ' 현재 카테고리 기준의 Task 로드
'    Dim tasksAll As Collection
'    Dim s As Date, e As Date
'    Dim cat As String: cat = CStr(Me.cmbTaskCategory.Value)
'    If Me.chkTaskShowAll.Value Then
'        Set tasksAll = LoadAllTasks_File_Cat(cat)                ' 전체
'    Else
'        GetVisibleDateRange s, e
'        'Set tasksAll = LoadTasksForDateRange_File_Cat(cat, s, e) ' 보이는 구간
'        Set tasksAll = LoadTasksForDateRange_File_Cat(s, e, cat) ' 보이는 구간
'    End If
'
'    ' 2) ON이면 새 Task로 다시 적용
'    If mTaskOverlay Then
'        ApplyTaskOverlay tasksAll
'    End If
'
'    ' 선택 From~To는 항상 가장 강하게 보여야 하므로 마지막에 다시 칠함
'    'PaintRange
'    ApplySelectedRangeOverlay
    
End Sub
'=== 선택연동 체크박스 ===
Private Sub chkTaskLinkSel_Click()
    mTaskLinkSel = (Me.chkTaskLinkSel.Value = True)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_LINKSEL, mTaskLinkSel
End Sub

'=== 리스트 선택 시 From~To 반영 ===
' ListBox 이벤트 연결
Private Sub lstTask_Click()
    HandleTaskListSelectionChanged
End Sub

Private Sub lstTask_Change()
    HandleTaskListSelectionChanged
End Sub

' frmYearCalendar 내에 추가
Private Sub lstTask_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH

    Dim r As Long
    r = Me.lstTask.ListIndex
    If r < 0 Then Exit Sub                      ' 선택 없음

    Dim sFrom As String
    sFrom = NzCStr(Me.lstTask.list(r, 1))       ' 0:Name, 1:From, 2:To
    If Len(Trim$(sFrom)) = 0 Then Exit Sub

    Dim dF As Date
    If TryParseYMD(sFrom, dF) Or TryParseDate(sFrom, dF) Then
        ' 연도 변경 + 화면 재렌더
        SetBaseYear Year(dF)
    Else
        MsgBox "선택 항목의 From 날짜를 해석할 수 없습니다: " & sFrom, vbExclamation
    End If
    Exit Sub
EH:
    ' 필요시 로깅 정도만
End Sub


'=== 날짜로 DayBox 찾기 (단일) ===
Private Function FindDayBoxByDate(ByVal d As Date) As MSForms.TextBox
    Dim key As String: key = Format$(d, "yyyy-mm-dd")
    If Not DayBoxByDate Is Nothing Then
        If DayBoxByDate.Exists(key) Then
            Set FindDayBoxByDate = DayBoxByDate(key)
        End If
    End If
End Function

'=== 오버레이 해제: 원래 테두리/툴팁 복원 (단일) ===
Private Sub ClearTaskOverlay()
    ClearAllDayBoxOverlay
    ' 선택 구간 하이라이트만 얹기
    ApplySelectedRangeOverlay
End Sub


'=== 오버레이 적용: 파란 테두리 + ToolTip 누적 (단일) ===
'=== 오버레이 적용: 테두리 + "배경 틴트" + ToolTip 누적 ===

Private Sub ApplyTaskOverlay(ByVal tasks As Collection)
    ' 기본/선택 상태로 초기화 + 테두리 제거
    ClearTaskOverlay
    If tasks Is Nothing Then Exit Sub

    Dim haveSel As Boolean, sSel As Date, eSel As Date
    GetSelectedRange haveSel, sSel, eSel

    Dim excl As Boolean: excl = (Me.chkExcludeNonBiz.Value = True)
    EnsureColorsForTasks tasks

    Dim i As Long, d As Date, d1 As Date, d2 As Date
    For i = 1 To tasks.Count
        Dim t As clsTaskItem, key As String, edge As Long, fill As Long
        Set t = tasks(i)
        key = MakeTaskKey(t)
        edge = ColorForTaskKey(key)
        fill = OverlayFillColorFor(edge)

        d1 = t.FromDate
        d2 = IIf(t.HasTo, t.ToDate, t.FromDate)

        For d = d1 To d2
            If excl And (IsWeekend(d) Or IsHoliday(d)) Then GoTo ContinueNextDate

            Dim tb As MSForms.TextBox
            Set tb = FindDayBoxByDate(d)
            If Not tb Is Nothing Then
                On Error Resume Next
                ' 테두리(한 번만)
                If tb.BorderStyle = 0 Then
                    tb.BorderStyle = fmBorderStyleSingle
                    tb.BorderColor = edge
                End If

                ' 선택 구간은 선택색이 최우선 → 배경은 바꾸지 않음
                If Not (haveSel And d >= sSel And d <= eSel And _
                        Not (excl And (IsWeekend(d) Or IsHoliday(d)))) Then
                    ' 오버레이 연한 배경 유지
                    tb.BackColor = fill
                End If

                ' 툴팁 누적
                Dim tip As String: tip = tb.ControlTipText
                If Len(t.TaskName) > 0 Then
                    If InStr(1, tip, t.TaskName, vbTextCompare) = 0 Then
                        tb.ControlTipText = IIf(Len(tip) = 0, t.TaskName, tip & vbCrLf & t.TaskName)
                    End If
                End If
                On Error GoTo 0
            End If
ContinueNextDate:
        Next d
    Next i
End Sub

' 폼 모듈 내 아무 곳
Private Sub ClearAllDayBoxBorders()
    Dim m As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox
    For m = LBound(tbDay, 1) To UBound(tbDay, 1)
        For r = LBound(tbDay, 2) To UBound(tbDay, 2)
            For c = LBound(tbDay, 3) To UBound(tbDay, 3)
                Set tb = tbDay(m, r, c)
                If Not tb Is Nothing Then
                    On Error Resume Next
                    tb.BorderStyle = 0     ' ★ 기본: 무테
                    On Error GoTo 0
                End If
            Next c
        Next r
    Next m
End Sub

'----- Task Color

Private Function TaskPalette() As Variant
    ' 눈에 잘 띄고 서로 구분되는 12색(필요시 더 추가)

    TaskPalette = Array( _
        RGB(66, 133, 244), _
        RGB(219, 68, 55), _
        RGB(244, 180, 0), _
        RGB(15, 157, 88), _
        RGB(171, 71, 188), _
        RGB(0, 172, 193), _
        RGB(255, 112, 67), _
        RGB(142, 36, 170), _
        RGB(85, 139, 47), _
        RGB(121, 134, 203), _
        RGB(233, 30, 99), _
        RGB(0, 121, 107) _
    )
End Function

' 플랫폼 무관, 오버플로우/형변환 이슈 없는 해시
' 반환: 0 .. 2,147,483,647  (31-bit 양수)
Private Function HashString(ByVal s As String) As Long
    Dim i As Long, ch As Long
    Dim h As Double: h = 5381#                 ' Double 누적
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))               ' 유니코드 코드포인트(16bit)
        If ch < 0 Then ch = ch + 65536         ' 무부호 16bit로 정규화
        h = h * 33# + ch                       ' DJB2 가산형 (비트연산 제거)
        ' 2^31 로 모듈러: 0 <= h < 2^31 이 되도록 접기
        h = h - 2147483648# * Fix(h / 2147483648#)
    Next
    HashString = CLng(h)                       ' 안전: 0..2147483647 범위
End Function

Private Function MakeTaskKey(ByVal t As clsTaskItem) As String
    ' 이름이 있으면 이름으로, 없으면 기간 문자열로 키 생성(색 일관성 유지 목적)
    Dim nm As String: nm = Trim$(t.TaskName)
    If Len(nm) > 0 Then
        MakeTaskKey = nm
    Else
        If t.HasTo Then
            MakeTaskKey = Format$(t.FromDate, "yyyy-mm-dd") & "~" & Format$(t.ToDate, "yyyy-mm-dd")
        Else
            MakeTaskKey = Format$(t.FromDate, "yyyy-mm-dd")
        End If
    End If
End Function

Private Function TryLoadTaskColorFromReg(ByVal key As String, ByRef clr As Long) As Boolean
    On Error GoTo EH
    Dim s As String
    s = GetSetting(REG_APP, REG_SEC_TASKCOLOR, key, "")
    If Len(s) = 0 Then Exit Function
    clr = CLng(s)           ' 저장을 그냥 Long 문자열로 했으니 그대로 복원
    TryLoadTaskColorFromReg = True
    Exit Function
EH:
    TryLoadTaskColorFromReg = False
End Function

Private Sub SaveTaskColorToReg(ByVal key As String, ByVal clr As Long)
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC_TASKCOLOR, key, CStr(clr)
End Sub

Private Function ColorForTaskKey(ByVal key As String) As Long
    ' 캐시→레지스트리→팔레트 해시 순서로 색을 결정
    If TaskColorByName Is Nothing Then Set TaskColorByName = CreateObject("Scripting.Dictionary")

    If TaskColorByName.Exists(key) Then
        ColorForTaskKey = TaskColorByName(key)
        Exit Function
    End If

    Dim c As Long
    If TryLoadTaskColorFromReg(key, c) Then
        TaskColorByName.add key, c
        ColorForTaskKey = c
        Exit Function
    End If

    ' 레지스트리에 없으면 고정 팔레트에서 해시로 선택
    Dim p As Variant: p = TaskPalette()
    Dim idx As Long: idx = HashString(UCase$(key)) Mod (UBound(p) - LBound(p) + 1)
    c = p(LBound(p) + idx)

    TaskColorByName.add key, c
    SaveTaskColorToReg key, c
    ColorForTaskKey = c
End Function

Private Sub EnsureColorsForTasks(ByVal tasks As Collection)
    ' 미리 한 번 돌려서 전부 색 준비(선택적이지만 가독성↑)
    If tasks Is Nothing Then Exit Sub
    Dim i As Long, k As String
    For i = 1 To tasks.Count
        k = MakeTaskKey(tasks(i))
        Call ColorForTaskKey(k)
    Next
End Sub

' lstTask에 동일 항목이 있으면 그 인덱스(0-base), 없으면 -1
Private Function FindTaskRow(ByVal lst As MSForms.ListBox, _
                             ByVal nameKey As String, _
                             ByVal fromYMD As String, _
                             ByVal toYMD As String) As Long
    Dim i As Long
    Dim nm As String, f As String, t As String

    For i = 0 To lst.ListCount - 1
        nm = NzCStr(lst.list(i, 0))
        f = NzCStr(lst.list(i, 1))
        t = NzCStr(lst.list(i, 2))

        If StrComp(nm, nameKey, vbTextCompare) = 0 _
           And f = fromYMD _
           And t = toYMD Then
            FindTaskRow = i
            Exit Function
        End If
    Next
    FindTaskRow = -1
End Function

Private Sub btnRemoveTaskYear_Click()
    Dim y As Long: y = GetCurrentYearForTask()
    If y < 1901 Or y > 9998 Then
        MsgBox "유효한 연도가 아닙니다: " & CStr(y), vbExclamation
        Exit Sub
    End If
    Dim cat As String: cat = SelectedCategoryName

    Dim resp As VbMsgBoxResult
    resp = MsgBox("카테고리 '" & cat & "'에서 연도 " & y & "의 Task를 모두 삭제합니다." & vbCrLf & _
                  "되돌릴 수 없습니다. 진행하시겠습니까?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "연도별 Task 삭제 확인")
    If resp <> vbYes Then Exit Sub

    On Error GoTo EH
    RemoveTasksForYear_FromAll_Cat cat, y
    RefreshTaskListAndOverlay
    MsgBox "연도 " & y & "의 Task가 삭제되었습니다.", vbInformation
    Exit Sub
EH:
    MsgBox "삭제 중 오류: " & Err.Description, vbExclamation
End Sub

Private Function GetRegBool(sec As String, key As String, defVal As Boolean) As Boolean
    Dim s As String
    s = GetSetting(REG_APP, sec, key, IIf(defVal, "1", "0"))
    GetRegBool = (s = "1")
End Function

Private Sub SaveRegBool(sec As String, key As String, v As Boolean)
    On Error Resume Next
    SaveSetting REG_APP, sec, key, IIf(v, "1", "0")
End Sub

Private Sub btnGetFromTo_Click()
    Dim sFrom As String, sTo As String
    Dim dF As Date, dT As Date

    sFrom = Trim$(Me.txtFromSel.text)
    sTo = Trim$(Me.txtToSel.text)

    ' From
    If Len(sFrom) > 0 Then
        If TryParseYMD(sFrom, dF) Or TryParseDate(sFrom, dF) Then
            Me.txtTaskFrom.text = Format$(dF, "yyyy-mm-dd")
        Else
            Me.txtTaskFrom.text = sFrom
        End If
    Else
        Me.txtTaskFrom.text = ""
    End If

    ' To
    If Len(sTo) > 0 Then
        If TryParseYMD(sTo, dT) Or TryParseDate(sTo, dT) Then
            Me.txtTaskTo.text = Format$(dT, "yyyy-mm-dd")
        Else
            Me.txtTaskTo.text = sTo
        End If
    Else
        Me.txtTaskTo.text = ""
    End If
End Sub
' 보이는 12개 월의 첫째날~마지막날
Private Sub GetVisibleDateRange(ByRef dStart As Date, ByRef dEnd As Date)
    Dim baseYear As Long: baseYear = mYear
    Dim curM As Long: curM = Month(Date)
    Dim mapDate() As Date
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot
    dStart = DateSerial(Year(mapDate(1)), Month(mapDate(1)), 1)
    dEnd = DateSerial(Year(mapDate(12)), Month(mapDate(12)) + 1, 0)
End Sub

' 보이는 달력 범위의 모든 연도에서 Task 로드 → 이름/기간 병합 → 리스트/오버레이 적용
Private Sub LoadTasksForVisibleRangeAndOverlay()
    Dim s As Date, e As Date
    GetVisibleDateRange s, e

    Dim col As Collection
    Set col = LoadTasksForDateRange_File(s, e)   ' ← JSON 파일에서 범위 로드 + 병합

    FillListBoxFromTasksSafe Me.lstTask, col

    BuildDayBoxMapFromGrid
    If mTaskOverlay Then
        ApplyTaskOverlay col
    Else
        ClearTaskOverlay
    End If
End Sub

' 대량 데이터를 빠르게 채움(.AddItem 루프 대신 2D 배열로 .List 지정)
Private Sub FillListBoxFromTasksFast(lst As MSForms.ListBox, ByVal tasks As Collection)
    Dim n As Long: n = IIf(tasks Is Nothing, 0, tasks.Count)
    Dim v As Variant
    Dim i As Long

    lst.Clear
    lst.ColumnCount = 3
    If n = 0 Then Exit Sub

    ReDim v(0 To n - 1, 0 To 2)
    For i = 1 To n
        Dim t As clsTaskItem
        Set t = tasks(i)
        v(i - 1, 0) = t.TaskName
        v(i - 1, 1) = Format$(t.FromDate, "yyyy-mm-dd")
        v(i - 1, 2) = IIf(t.HasTo, Format$(t.ToDate, "yyyy-mm-dd"), "")
    Next
    lst.list = v
End Sub

Private Sub chkTaskShowAll_Click()
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_SHOWALL, (Me.chkTaskShowAll.Value = True)
    RefreshTaskListAndOverlay
End Sub

Private Sub RefreshTaskListAndOverlay()
    Dim cat As String: cat = CStr(Me.cmbTaskCategory.Value)
    Dim tasks As Collection
    Dim displayTasks As Collection
    Dim tasksForOverlay As Collection
    Dim s As Date, e As Date

    ' ① 로드 (표시용)
    If Me.chkTaskShowAll.Value Then
        Set tasks = LoadAllTasks_File_Cat(cat)                  ' 표시용(전체)
    Else
        GetVisibleDateRange s, e
        Set tasks = LoadTasksForDateRange_File_Cat(s, e, cat)   ' 표시용(구간)
    End If

    ' 리스트에는 이름 필터(txtTaskFilter)만 적용
    Set displayTasks = FilterTasksByNameLike(tasks, NzCStr(Me.txtTaskFilter.text))

    FillListBoxFromTasksSafe Me.lstTask, displayTasks  ' (필터 쓰는 경우 기존대로 변경)

    ' ③ 오버레이는 "모든 Task" 기준으로 (요건 유지)
    '     - ShowAll이면 tasks 그대로, 아니면 범위 버전(필요 시 모든 연도 로드로 바꿔도 무방)
    Set tasksForOverlay = displayTasks

    ' ④ DayBox 매핑 최신화
    BuildDayBoxMapFromGrid
    
    ' 핵심: 먼저 전부 베이스로 초기화(툴팁에 Task 이름 제거 포함)
    ClearAllDayBoxOverlay

    ' ⑥ 오버레이가 켜져 있으면 이번 카테고리 Task로 다시 적용
    If mTaskOverlay Then
        ApplyTaskOverlay tasksForOverlay
    End If

    ' ⑦ 선택 From~To가 항상 최상위로 보이도록 마지막에 다시 칠함
    'PaintRange
    ApplySelectedRangeOverlay
    
End Sub

'Private Sub txtTaskFilter_Change()
'    RefreshTaskListAndOverlay
'End Sub

Private Sub btnTaskFilterApply_Click()
    RefreshTaskListAndOverlay
End Sub

Private Sub btnTaskLoad_Click()
    RefreshTaskListAndOverlay
End Sub

' 모든 열을 행별로 명시 대입 (IIf 사용 안함)
Private Sub FillListBoxFromTasksSafe(lst As MSForms.ListBox, ByVal tasks As Collection)
    Dim n As Long: n = IIf(tasks Is Nothing, 0, tasks.Count)
    Dim data As Variant
    Dim i As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnHeads = False
    ' 필요 시 가독성: lst.ColumnWidths = "140 pt;75 pt;75 pt"
    lst.BoundColumn = 0

    If n = 0 Then Exit Sub

    ReDim data(0 To n - 1, 0 To 2)
    For i = 1 To n
        Dim t As clsTaskItem
        Set t = tasks(i)

        data(i - 1, 0) = t.TaskName                     ' Name
        data(i - 1, 1) = Format$(t.FromDate, "yyyy-mm-dd") ' From

        If t.HasTo Then
            data(i - 1, 2) = Format$(t.ToDate, "yyyy-mm-dd")
        Else
            data(i - 1, 2) = ""
        End If
    Next

    lst.list = data
End Sub

Private Function TaskEndDate(ByVal t As clsTaskItem) As Date
    If t.HasTo Then TaskEndDate = t.ToDate Else TaskEndDate = t.FromDate
End Function

' 선택 항목과 동일한 (Name, From, To) Task를 JSON 저장소에서 삭제
Private Sub btnTaskDeleteSelected_Click()
    Dim sel As Collection: Set sel = SelectedTasksFromList(Me.lstTask)
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "삭제할 항목을 선택하세요.", vbExclamation
        Exit Sub
    End If

    Dim cat As String: cat = SelectedCategoryName
    Dim all As Collection: Set all = LoadAllTasks_File_Cat(cat)

    ' 선택 키 집합
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    Dim i As Long, t As clsTaskItem
    For i = 1 To sel.Count
        keys(TripleKey(sel(i))) = True
    Next

    ' 필터링
    Dim remain As New Collection
    For i = 1 To all.Count
        Set t = all(i)
        If Not keys.Exists(TripleKey(t)) Then remain.add t
    Next

    On Error GoTo EH
    SaveAllTasks_File_Cat cat, remain
    RefreshTaskListAndOverlay
    MsgBox "선택 항목을 삭제했습니다.", vbInformation
    Exit Sub
EH:
    MsgBox "삭제 중 오류: " & Err.Description, vbExclamation
End Sub

' name|yyyy-mm-dd|yyyy-mm-dd(또는 빈문자열) 로 비교키 생성
Private Function TripleKey(ByVal t As clsTaskItem) As String
    Dim f As String, z As String
    f = Format$(t.FromDate, "yyyy-mm-dd")
    z = IIf(t.HasTo, Format$(t.ToDate, "yyyy-mm-dd"), "")
    TripleKey = NzCStr(t.TaskName) & "|" & f & "|" & z
End Function

' (name, from, to) → 비교용 키 (이름은 대소문자 무시)
Private Function MakeTripleKey(ByVal nm As String, ByVal fromYMD As String, ByVal toYMD As String) As String
    MakeTripleKey = LCase$(Trim$(NzCStr(nm))) & "|" & Trim$(NzCStr(fromYMD)) & "|" & Trim$(NzCStr(toYMD))
End Function

' clsTaskItem → 비교용 키
Private Function TaskTripleKey(ByVal t As clsTaskItem) As String
    TaskTripleKey = MakeTripleKey( _
        t.TaskName, _
        Format$(t.FromDate, "yyyy-mm-dd"), _
        IIf(t.HasTo, Format$(t.ToDate, "yyyy-mm-dd"), "") _
    )
End Function

' lstTask에서 선택된 모든 항목의 키 집합(Dictionary) 반환
Private Function SelectedTripleKeysFromList(lst As MSForms.ListBox) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            Dim nm As String, f As String, t As String
            nm = NzCStr(lst.list(i, 0))
            f = NzCStr(lst.list(i, 1))
            t = NzCStr(lst.list(i, 2))
            d(MakeTripleKey(nm, f, t)) = True
        End If
    Next
    Set SelectedTripleKeysFromList = d
End Function


' ListBox의 선택 변경 시 호출: 선택된 Task만 오버레이 + From~To = Min~Max
Private Sub HandleTaskListSelectionChanged()

    ' === 링크 연동이 꺼져 있으면 아무 것도 안 함 ===
    If Not (Me.chkTaskLinkSel.Value = True) Then Exit Sub
    
    Dim sel As Collection: Set sel = SelectedTasksFromList(Me.lstTask)

    If sel Is Nothing Or sel.Count = 0 Then
        ' 선택 없을 때는 기존 단일 From~To 로직으로 복원
        Set mSelRanges = Nothing
        PaintRange                      ' 기존 단일 범위 칠하기
        ' 오버레이는 유지됨
        Exit Sub
    End If

    ' 1) 여러 구간 계산
    Set mSelRanges = BuildRangesFromTasks(sel)

    ' 2) 최솟값/최댓값으로 From~To 텍스트 및 내부 상태 갱신
    Dim i As Long, s As Date, e As Date
    Dim minF As Date, maxT As Date
    minF = sel(1).FromDate: maxT = TaskEndDate(sel(1))
    For i = 2 To sel.Count
        s = sel(i).FromDate
        e = TaskEndDate(sel(i))
        If s < minF Then minF = s
        If e > maxT Then maxT = e
    Next

    mFrom = minF: mHasFrom = True
    mTo = maxT: mHasTo = True
    Me.txtFromSel.text = FmtDateOut(minF)
    Me.txtToSel.text = FmtDateOut(maxT)

    ' 3) 선택된 구간들만 배경 칠하기 (오버레이 테두리는 건드리지 않음)
    PaintSelectionRangesMulti mSelRanges

    ' (선택) 오버레이가 켜져 있다면 테두리를 최신 화면에 재적용하고 싶을 때:
    If mTaskOverlay Then RefreshTaskOverlayIfOn
    
    Dim selCount As Long, firstIdx As Long
    firstIdx = -1

    For i = 0 To Me.lstTask.ListCount - 1
        If Me.lstTask.Selected(i) Then
            selCount = selCount + 1
            If firstIdx < 0 Then firstIdx = i
            If selCount > 1 Then Exit For  ' 다중 선택이면 텍스트박스 세팅 안 함
        End If
    Next

    If selCount = 1 And firstIdx >= 0 Then
        With Me.lstTask
            Me.txtTaskName.text = NzCStr(.list(firstIdx, 0))
            Me.txtTaskFrom.text = NzCStr(.list(firstIdx, 1))
            Me.txtTaskTo.text = NzCStr(.list(firstIdx, 2))
        End With
    End If
    
End Sub

' 선택된 행들만 clsTaskItem으로 변환
Private Function SelectedTasksFromList(lst As MSForms.ListBox) As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            Dim t As clsTaskItem
            Set t = TaskFromRow(lst, i)
            If Not t Is Nothing Then out.add t
        End If
    Next
    Set SelectedTasksFromList = out
End Function

' 특정 행을 clsTaskItem으로 변환
Private Function TaskFromRow(lst As MSForms.ListBox, ByVal row As Long) As clsTaskItem
    Dim nm As String, sf As String, st As String
    Dim dF As Date, dT As Date, okF As Boolean, okT As Boolean

    nm = NzCStr(lst.list(row, 0))
    sf = NzCStr(lst.list(row, 1))
    st = NzCStr(lst.list(row, 2))

    okF = (TryParseYMD(sf, dF) Or TryParseDate(sf, dF))
    If Not okF Then Exit Function

    Dim t As clsTaskItem
    Set t = New clsTaskItem
    t.TaskName = nm
    t.FromDate = dF

    okT = (Len(st) > 0) And (TryParseYMD(st, dT) Or TryParseDate(st, dT))
    t.HasTo = okT
    If okT Then t.ToDate = dT

    Set TaskFromRow = t
End Function

' 선택된 Task들로부터 [시작~끝] 구간 배열 생성
Private Function BuildRangesFromTasks(ByVal tasks As Collection) As Collection
    Dim rngs As New Collection, i As Long, s As Date, e As Date, tmp As Date
    If Not tasks Is Nothing Then
        For i = 1 To tasks.Count
            s = tasks(i).FromDate
            e = IIf(tasks(i).HasTo, tasks(i).ToDate, s)
            If e < s Then tmp = s: s = e: e = tmp
            rngs.add Array(s, e)   ' Variant(2) 사용
        Next
    End If
    Set BuildRangesFromTasks = rngs
End Function

' 날짜가 임의의 구간들 안에 포함되는가?
Private Function DateInRanges(ByVal d As Date, ByVal ranges As Collection) As Boolean
    Dim i As Long, a As Variant, s As Date, e As Date
    For i = 1 To ranges.Count
        a = ranges(i): s = a(0): e = a(1)
        If d >= s And d <= e Then
            DateInRanges = True
            Exit Function
        End If
    Next
    DateInRanges = False
End Function

' 여러 구간을 “선택영역” 색으로 칠하기 (Border는 건드리지 않음 → 오버레이 유지)
Private Sub PaintSelectionRangesMulti(ByVal ranges As Collection)
    Dim mBlock As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox, sTag As String, dT As Date
    Dim excludeNonBiz As Boolean: excludeNonBiz = chkExcludeNonBiz.Value

    For mBlock = 1 To 12
        For r = 1 To 6
            For c = 1 To 7
                Set tb = tbDay(mBlock, r, c)
                sTag = Trim$(tb.Tag)

                ' 기본 스타일 복구 (배경/전경/오늘/휴일 등) - Border는 만지지 않음
                ApplyBaseStyle tb, c

                If Len(sTag) > 0 And TryParseYMD(sTag, dT) Then
                    If DateInRanges(dT, ranges) Then
                        If excludeNonBiz Then
                            If Not (IsWeekend(dT) Or IsHoliday(dT)) Then
                                'tb.BackColor = RGB(255, 244, 204)
                                tb.BackColor = RGB(33, 92, 152)
                                tb.ForeColor = RGB(255, 255, 255)
                                tb.Font.Bold = True
                            End If
                        Else
                            'tb.BackColor = RGB(255, 244, 204)
                            tb.BackColor = RGB(33, 92, 152)
                            tb.ForeColor = RGB(255, 255, 255)
                            tb.Font.Bold = True
                        End If
                    End If
                End If
            Next
        Next
    Next
    UpdateRangeInfo
End Sub

Private Sub LoadCategoryList()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fld As Object, f As Object, nm As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' 기본 카테고리 "tasks"는 항상 포함
    dict.add "tasks", True

    EnsureFolder TaskDataRootFolder
    Set fld = fso.GetFolder(TaskDataRootFolder)
    For Each f In fld.Files
        If LCase$(fso.GetExtensionName(f.path)) = "json" Then
            nm = fso.GetBaseName(f.path)
            If Len(nm) > 0 Then dict(nm) = True
        End If
    Next

    Me.cmbTaskCategory.Clear
    Dim k As Variant
    For Each k In dict.keys
        Me.cmbTaskCategory.AddItem CStr(k)
    Next

    ' 마지막 선택 복원(없으면 tasks)
    Dim last As String
    last = GetSetting(REG_APP, REG_SEC_TASKPANEL, REG_KEY_TASK_CATEGORY, "tasks")
    If Not SelectComboValue(Me.cmbTaskCategory, last) Then
        SelectComboValue Me.cmbTaskCategory, "tasks"
    End If
End Sub

Private Function SelectComboValue(cb As MSForms.ComboBox, ByVal v As String) As Boolean
    Dim i As Long
    For i = 0 To cb.ListCount - 1
        If StrComp(cb.list(i), v, vbTextCompare) = 0 Then
            cb.ListIndex = i
            SelectComboValue = True
            Exit Function
        End If
    Next
    SelectComboValue = False
End Function

Private Function SelectedCategoryName() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(Me.cmbTaskCategory.Value)
    On Error GoTo 0
    If Len(s) = 0 Then s = "tasks"
    SelectedCategoryName = s
End Function

Private Sub cmbTaskCategory_Change()
    SaveSetting REG_APP, REG_SEC_TASKPANEL, REG_KEY_TASK_CATEGORY, SelectedCategoryName
    txtTaskName.text = "": txtTaskFrom.text = "": txtTaskTo.text = ""
    RefreshTaskListAndOverlay    ' 카테고리 바뀌면 즉시 다시 로드/오버레이
    ApplySelectedRangeOverlay
End Sub

Private Sub cmbTaskCategory_Click()
    SaveSetting REG_APP, REG_SEC_TASKPANEL, REG_KEY_TASK_CATEGORY, SelectedCategoryName
    txtTaskName.text = "": txtTaskFrom.text = "": txtTaskTo.text = ""
'    RefreshTaskListAndOverlay
'    ApplySelectedRangeOverlay
    ReapplyOverlayAndSelection (True)
End Sub

Private Sub btnTaskCategoryAdd_Click()
    Dim nm As String
    nm = InputBox("추가할 유형명을 입력하세요.", "유형 추가")
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Sub

    nm = SafeFileBaseName(nm)

    ' 이미 목록에 있으면 선택만
    If SelectComboValue(Me.cmbTaskCategory, nm) Then Exit Sub

    ' 빈 파일 생성([]) + 로그 스냅샷
    Dim emptyCol As New Collection
    SaveAllTasks_File_Cat nm, emptyCol

    ' 콤보 갱신 및 선택
    LoadCategoryList
    SelectComboValue Me.cmbTaskCategory, nm
End Sub

Private Sub btnTaskCategoryDelete_Click()
    Dim cat As String: cat = SelectedCategoryName
    If LCase$(cat) = "tasks" Then
        MsgBox "기본 카테고리 'tasks'는 삭제할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Dim r As VbMsgBoxResult
    r = MsgBox("유형 '" & cat & "'의 JSON 파일을 삭제합니다." & vbCrLf & _
               "로그 폴더의 해당 스냅샷도 함께 삭제할까요?", _
               vbQuestion + vbYesNoCancel, "유형 삭제")
    If r = vbCancel Then Exit Sub

    On Error Resume Next
    ' 본 파일 삭제
    RemoveAllTasks_File_Cat cat

    ' 로그 삭제 선택(vbYes면 로그도 삭제)
    If r = vbYes Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        Dim fld As Object, f As Object, base As String
        base = SafeFileBaseName(cat) & "_"
        EnsureFolder TaskLogFolder
        Set fld = fso.GetFolder(TaskLogFolder)
        For Each f In fld.Files
            If LCase$(Left$(fso.GetBaseName(f.path), Len(base))) = LCase$(base) Then
                f.Delete True
            End If
        Next
    End If
    On Error GoTo 0

    LoadCategoryList
    RefreshTaskListAndOverlay
End Sub


' name이 filterRaw(예: "dev,release*") 중 하나라도 매칭되면 True
Private Function MatchesAnyFilter(ByVal name As String, ByVal filterRaw As String) As Boolean
    Dim raw As String, parts() As String, i As Long, pat As String
    raw = Trim$(filterRaw)
    If raw = "" Or raw = "*" Then
        MatchesAnyFilter = True
        Exit Function
    End If

    parts = Split(raw, ",")
    For i = LBound(parts) To UBound(parts)
        pat = Trim$(parts(i))
        If Len(pat) > 0 Then
            ' 와일드카드가 없으면 부분일치로 변환
            If InStr(1, pat, "*") = 0 And InStr(1, pat, "?") = 0 Then
                pat = "*" & pat & "*"
            End If
            If LCase$(name) Like LCase$(pat) Then
                MatchesAnyFilter = True
                Exit Function
            End If
        End If
    Next
End Function

' 컬렉션에서 TaskName 기준으로 필터
Private Function FilterTasksByNameLike(ByVal tasks As Collection, ByVal filterRaw As String) As Collection
    Dim out As New Collection, i As Long, t As clsTaskItem, nm As String
    Dim raw As String: raw = Trim$(filterRaw)
    If raw = "" Or raw = "*" Then
        ' 필터 비활성: 원본 그대로 반환(참조 반환이라 성능↑)
        Set FilterTasksByNameLike = tasks
        Exit Function
    End If

    For i = 1 To IIf(tasks Is Nothing, 0, tasks.Count)
        Set t = tasks(i)
        nm = NzCStr(t.TaskName)
        If MatchesAnyFilter(nm, raw) Then out.add t
    Next
    Set FilterTasksByNameLike = out
End Function

' 폼 모듈(frmYearCalendar)
Private Function CurrentCategoryName() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(Me.cmbTaskCategory.Value)
    On Error GoTo 0
    If Len(s) = 0 Then s = "tasks"   ' 기본 카테고리
    CurrentCategoryName = s
End Function

Private Function LoadTasksForDateRange_File_Cat(ByVal dStart As Date, _
                                                ByVal dEnd As Date, _
                                                ByVal cat As String) As Collection
    Dim all As Collection: Set all = LoadAllTasks_File_Cat(cat)
    Dim filtered As New Collection
    Dim i As Long, t As clsTaskItem, s As Date, e As Date
    For i = 1 To all.Count
        Set t = all(i)
        s = t.FromDate
        e = TaskEndDate(t)
        If IntersectsRange(s, e, dStart, dEnd) Then filtered.add t
    Next
    'Set LoadTasksForDateRange_File_Cat = MergeTasksByNameAndAdjacency(filtered)
    Set LoadTasksForDateRange_File_Cat = filtered
End Function

' === 색 도우미 ===
Private Function MixColors(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    ' 결과 = c1*(1-t) + c2*t  (t: 0~1)
    If t < 0 Then
        t = 0
    Else
        If t > 1 Then
            t = 1
        End If
    End If
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = (c1 And &HFF&): g1 = (c1 \ &H100 And &HFF&): b1 = (c1 \ &H10000 And &HFF&)
    r2 = (c2 And &HFF&): g2 = (c2 \ &H100 And &HFF&): b2 = (c2 \ &H10000 And &HFF&)
    MixColors = RGB( _
        CLng(r1 + (r2 - r1) * t), _
        CLng(g1 + (g2 - g1) * t), _
        CLng(b1 + (b2 - b1) * t))
End Function

Private Function OverlayFillColorFor(ByVal edgeColor As Long) As Long
    ' 테두리색을 흰색 쪽으로 70% 밝게 → 은은한 칠(오버레이 영역)
    OverlayFillColorFor = MixColors(edgeColor, RGB(255, 255, 255), 0.7)
End Function


' 폼 모듈(frmYearCalendar) 어딘가에 추가
Private Sub PaintSelectedCell(ByVal tb As MSForms.TextBox)
    With tb
        .BackColor = RGB(33, 92, 152)   ' 진한 파랑
        .ForeColor = RGB(255, 255, 255) ' 흰 글자
        .Font.Bold = True
        '원하면 테두리도 통일
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(33, 92, 152)
    End With
End Sub

'=== REPLACE: 오버레이/툴팁 완전 초기화(주말/공휴일/오늘 색 복원 + 테두리 끄기 + 툴팁=날짜/공휴일만) ===
Private Sub ClearAllDayBoxOverlay()
    Dim m As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox
    Dim sTag As String, d As Date

    For m = LBound(tbDay, 1) To UBound(tbDay, 1)
        For r = LBound(tbDay, 2) To UBound(tbDay, 2)
            For c = LBound(tbDay, 3) To UBound(tbDay, 3)
                Set tb = tbDay(m, r, c)
                If Not tb Is Nothing Then
                    ' 1) 기본 색·글꼴 복구
                    ApplyBaseStyle tb, c
                    ' 2) 테두리 제거
                    On Error Resume Next
                    tb.BorderStyle = 0
                    On Error GoTo 0
                    ' 3) 툴팁을 '날짜(+공휴일)'만으로 재설정
                    sTag = Trim$(tb.Tag)
                    If Len(sTag) = 10 And TryParseYMD(sTag, d) Then
                        tb.ControlTipText = BaseTipForDate(d)
                    Else
                        tb.ControlTipText = ""
                    End If
                End If
            Next c
        Next r
    Next m
End Sub

'=== NEW: 날짜 + (있으면) 공휴일만으로 베이스 툴팁 생성 ===
Private Function BaseTipForDate(ByVal d As Date) As String
    Dim tip As String, holName As String, holPrefix As String
    tip = FmtDateOut(d)                                ' 날짜 1줄
    holName = GetHolidayNameIfAny(d)                   ' 공휴일 이름
    If Len(holName) > 0 Then
        holPrefix = IIf(mLang = LangE, "[Holiday] ", "[공휴일] ")
        tip = tip & vbCrLf & holPrefix & holName
    End If
    BaseTipForDate = tip
End Function

' 선택 구간을 항상 가장 마지막에 “덮어 칠” 하이라이트
Private Sub ApplySelectedRangeOverlay()
    Static mInSelOverlay As Boolean
    If mInSelOverlay Then Exit Sub
    mInSelOverlay = True
    On Error GoTo X

    Dim have As Boolean, s As Date, e As Date, d As Date
    GetSelectedRange have, s, e
    If Not have Then GoTo X

    Dim excludeNonBiz As Boolean
    excludeNonBiz = (Me.chkExcludeNonBiz.Value = True)

    For d = s To e
        ' ★ 제외 옵션이 켜져 있으면 토/일/공휴일은 건너뜀
        If excludeNonBiz Then
            If IsWeekend(d) Or IsHoliday(d) Then GoTo ContinueNextDate
        End If

        Dim tb As MSForms.TextBox
        Set tb = FindDayBoxByDate(d)
        If Not tb Is Nothing Then
            ' 선택 구간은 최우선(오버레이/주말색 등 위에 덮어씀)
            tb.BackColor = RGB(33, 92, 152)
            tb.ForeColor = RGB(255, 255, 255)
            tb.Font.Bold = True
        End If
ContinueNextDate:
    Next

X:
    mInSelOverlay = False
End Sub

Private Sub RefreshVisuals(Optional ByVal rerenderMonths As Boolean = False)
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo LFinally

    If rerenderMonths Then
        RenderAllMonths              ' 달력 셀 생성/기본 스타일
        BuildDayBoxMapFromGrid
    End If

    ClearTaskOverlay                ' 테두리/툴팁/배경 원복(선택 표시 X)
    If mTaskOverlay Then
        ApplyTaskOverlay TasksFromListBox(Me.lstTask)  ' 배경/테두리/툴팁
    End If

    PaintRange                      ' ★ 마지막에 선택범위 칠하기(가장 강하게)

LFinally:
    mRefreshing = False
End Sub

Private Sub ReapplyOverlayAndSelection(Optional ByVal rerender As Boolean = False)
    RefreshVisuals rerender                               ' §3 오케스트레이터
End Sub

Private Function SortTasksByFromDate(ByVal tasks As Collection) As Collection
    ' 단순 안정정렬: 배열로 옮겨 정렬 후 재수집
    Dim n As Long: n = tasks.Count
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        arr(i, 1) = tasks(i).FromDate
        Set arr(i, 2) = tasks(i)
    Next

    ' QuickSort by arr(*,1)
    QuickSort2D arr, 1, n

    Dim col As New Collection
    For i = 1 To n
        col.add arr(i, 2)
    Next
    Set SortTasksByFromDate = col
End Function

Private Sub QuickSort2D(ByRef a() As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim p As Variant
    i = lo: j = hi
    p = a((lo + hi) \ 2, 1)
    Do While i <= j
        Do While a(i, 1) < p: i = i + 1: Loop
        Do While a(j, 1) > p: j = j - 1: Loop
        If i <= j Then
            Dim t1 As Variant, t2 As Variant
            t1 = a(i, 1): a(i, 1) = a(j, 1): a(j, 1) = t1
            Set t2 = a(i, 2): Set a(i, 2) = a(j, 2): Set a(j, 2) = t2
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort2D a, lo, j
    If i < hi Then QuickSort2D a, i, hi
End Sub

Private Function TasksFromListBox(lst As MSForms.ListBox) As Collection
    Dim out As New Collection
    Dim i As Long
    Dim t As clsTaskItem

    For i = 0 To lst.ListCount - 1
        Set t = New clsTaskItem                 ' ★★ 매 반복마다 New 필수 ★★
        
        Dim sName As String, sFrom As String, sTo As String
        Dim dF As Date, dT As Date
        Dim okF As Boolean, okT As Boolean

        sName = NzCStr(lst.list(i, 0))
        sFrom = NzCStr(lst.list(i, 1))
        sTo = NzCStr(lst.list(i, 2))

        okF = (TryParseYMD(sFrom, dF) Or TryParseDate(sFrom, dF))
        okT = (Len(sTo) > 0) And (TryParseYMD(sTo, dT) Or TryParseDate(sTo, dT))

        If okF Then
            t.TaskName = sName
            t.FromDate = dF
            t.HasTo = okT
            If okT Then t.ToDate = dT
            out.add t                              ' ★ 새 인스턴스를 추가
        End If
    Next

    Set TasksFromListBox = out
End Function

