VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYearCalendar 
   Caption         =   "Year Calendar (Period Picker)  �� Author :  gninraw7@naver.com"
   ClientHeight    =   9416.001
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   19140
   OleObjectBlob   =   "frmYearCalendar.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmYearCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' �� 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit

' ���� ����ȭ �� ��� ���� �÷���
Private mSyncingYear As Boolean

' === Week-start & Weekname style ===
Private Const REG_KEY_WEEK_START As String = "WeekStart"         ' "Sun" / "Mon"
Private Const REG_KEY_WEEKNAME_STYLE As String = "WeekNameStyle" ' "Short" / "Full"

Private Enum eWeekStart
    WeekSun = 0   ' �� ����: �Ͽ���
    WeekMon = 1   ' �� ����: ������
End Enum

Private Enum eWeekNameStyle
    WkShort = 0   ' Sun, Mon, ...
    WkFull = 1    ' Sunday, Monday, ...
End Enum

Private mWeekStart As eWeekStart
Private mWeekNameStyle As eWeekNameStyle


' ===== ��� ���� =====
Private Enum eLang
    LangK = 0   ' Korean
    LangE = 1   ' English
End Enum

Private mLang As eLang
Private mFmtDate As String         ' ��: "yyyy-mm-dd" �Ǵ� "mmm d, yyyy"


' �⺻��(���� ���� ��)
Private Const DEF_FMT_DATE_K As String = "yyyy-mm-dd"
Private Const DEF_FMT_DATE_E As String = "yyyy-mm-dd"     ' �ʿ� �� "mmm d, yyyy" �� �ٲ㾲�� ��
Private Const DEF_FMT_TITLE_K As String = "yyyy""��"" m""��"""
Private Const DEF_FMT_TITLE_E As String = "mmmm yyyy"

' ========= Registry I/O & ��ƿ =========
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
Private Const REG_KEY_FMT_DATE As String = "Date"        ' txtFromSel/ToSel, ��ȯ, �� NumberFormat

' === Month title formats per language ===
Private Const REG_KEY_FMT_TITLE As String = "MonthTitle" ' �� Ÿ��Ʋ ��/��Ʈ Ÿ��Ʋ
Private Const REG_KEY_FMT_TITLE_K As String = "MonthTitle_K"
Private Const REG_KEY_FMT_TITLE_E As String = "MonthTitle_E"

Private mFmtMonthTitle As String      ' ���� �� ���� ���õǾ� ���
Private mFmtMonthTitleK As String     ' �ѱ� ��� ����
Private mFmtMonthTitleE As String     ' ���� ��� ����


' ==== Ű ����ũ ====
Private Const SHIFT_MASK As Integer = 1
Private Const CTRL_MASK  As Integer = 2
Private Const ALT_MASK   As Integer = 4

' === ��ġ ��� Ȯ�� ===
Private Enum eLayoutMode
    lmNormal = 0          ' 1��~12��
    lmCurrentFirst = 1    ' ������� 1��1��
    lmCurrentLast = 2     ' ������� 3��4��
    lmCurrentAtSlot = 3   ' ��������� ���� ����(1..12)�� ��ġ
End Enum

Private mLayoutMode As eLayoutMode

Private Const VK_OEM_PERIOD     As Long = 190
Private Const VK_OEM_COMMA      As Long = 188

' === ����� ���� ��Ŀ(1..12) ===
Private mAnchorSlot As Long  ' optCurrentAtSlot ��忡���� ���

' === �� �ڽ� ��ġ ��/�� (4x3) ===
Private Const GRID_COLS As Long = 4
Private Const GRID_ROWS As Long = 3

' === �߰�: ����� ��� �󺧵� ===
Private lblMonthBG(1 To 12) As MSForms.Label

' ��/����
Private mCurMonthBG As Long
Private mCurMonthBorder As Long
Private mBGPad As Single

' ========= ���� (Const ����: RGB�� �Լ��� ��Ÿ�ӿ� ����) =========
Private mClrBG As Long
Private mClrSunFg As Long, mClrSatFg As Long, mClrWeekFg As Long
Private mClrHolBg As Long, mClrTodayBg As Long, mClrRangeBg As Long
Private mClrMonthTitle As Long
Private mClrCurMonthBg As Long   ' ����� ���� ���

' === Month label ���� ���� �� ===
Private mYearOddBG  As Long   ' Ȧ�� ���� ���
Private mYearEvenBG As Long   ' ¦�� ���� ���
Private mYearFG     As Long   ' ���� ���� (���ڻ�)

' ========= ��ȯ Ÿ�� =========
Private Enum ePushMode
    PushNone = 0
    PushToRange = 1          ' ��Ʈ ����(Selection �Ǵ� ���� Range)
    PushToTextBoxes = 2      ' �ٸ� ���� TextBox �� ��(From/To)
End Enum

Private mPushMode As ePushMode
Private mTargetRange As Range
Private mTargetTextFrom As MSForms.TextBox
Private mTargetTextTo As MSForms.TextBox

' ========= ��Ʈ/���̾ƿ� =========
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

' ========= ���� =========
Private mYear As Long
Private mCreated As Boolean

' ���� ��Ʈ�� �迭
Private lblMonth(1 To 12) As MSForms.Label
Private lblWeek(1 To 12, 1 To 7) As MSForms.Label
Private tbDay(1 To 12, 1 To 6, 1 To 7) As MSForms.TextBox

' �̺�Ʈ �� ����
Private mHooks As Collection

' From/To ����
Private mHasFrom As Boolean, mHasTo As Boolean
Private mFrom As Date, mTo As Date

' ������ ĳ��
Private HolidayByDate As Object       ' key "yyyy-mm-dd" �� name
Private HolidayYearLoaded As Object   ' key "2025" �� True

'========= 2025.09.20 Task Panel =========
' Task Panel ��� ������
Private mBaseWidth As Single
Private mTaskWidth As Single
Private mTaskVisible As Boolean

' �� ���� ���: ���� -> TextBox ���� (�������̿�)
Private DayBoxByDate As Object ' Scripting.Dictionary
Private OrigBorderColor As Object     ' "yyyy-mm-dd" -> Long
Private OrigToolTip As Object         ' "yyyy-mm-dd" -> String

' ����� ��ȣ: Task ��������/���ÿ���
Private mTaskOverlay As Boolean
Private mTaskLinkSel As Boolean

'=== Task Color ���� ===
Private TaskColorByName As Object  ' key: TaskKey(���� TaskName), val: Long(OLE Color)
Private Const REG_SEC_TASKCOLOR As String = "TaskColors"  ' PeriodPicker\TaskColors �ؿ� ����
Private Const SEC_TASKS As String = "Tasks"      ' SaveSetting/GetSetting ����

'=== Task Panel prefs (Registry) ===
Private Const REG_SEC_TASKPANEL As String = "TaskPanel"
Private Const REG_KEY_TASK_VISIBLE As String = "Visible"   ' "1" / "0"
Private Const REG_KEY_TASK_OVERLAY As String = "Overlay"   ' "1" / "0"
Private Const REG_KEY_TASK_LINKSEL As String = "LinkSel"   ' "1" / "0"
Private Const REG_KEY_SHOWALL As String = "ShowAll"
Private Const REG_KEY_TASK_CATEGORY As String = "Category"   ' ������ ���� ī�װ� ����

' ���õ� Task���� ���� ����(���� ��)�� ĥ�ϱ� ���� ������
Private mSelRanges As Collection   ' �� ������: Variant(2) = Array(DateStart, DateEnd)

' === Overlay ��� ���� �� ��ħ ���� ���� ===
Private OrigBackColor As Object    ' key: "yyyy-mm-dd" -> Long(���� BackColor)
Private OverlayHitCount As Object  ' key: "yyyy-mm-dd" -> Long(��ħ Ƚ��)

' ���� ���̶���Ʈ ������ ����
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
    mClrCurMonthBg = RGB(240, 248, 255)  ' ������ ���̶���Ʈ(AliceBlue �迭)
    
    ' ���� ���� ���/���ڻ� (���ϴ� ������ ���� ����)
    mYearOddBG = RGB(255, 248, 230)    ' Ȧ����: �ణ ����
    mYearEvenBG = RGB(236, 244, 255)   ' ¦����: �ణ ����
    mYearFG = RGB(30, 30, 30)
    
    mCurMonthBG = RGB(255, 247, 205)     ' ������ ���
    mCurMonthBorder = RGB(240, 170, 0)   ' �׵θ� ��
    mBGPad = 2                           ' ��� �ٱ��� ����(px)
    
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

    ' ���� ���� �ݿ�
    mLayoutMode = m
    mAnchorSlot = slot

    ' UI �ݿ�
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
    
    ' Auto next row (apply �� ���ÿ����� �Ʒ��� �̵�)
    Dim sAuto As String
    sAuto = GetSetting(REG_APP, REG_SEC_LAYOUT, REG_KEY_AUTO_NEXTROW, "0")
    Me.chkAutoNextRow.Value = (sAuto = "1")
    
    
End Sub

' === ���� ===
Private Sub SaveLayoutPrefs()
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_MODE, CStr(mLayoutMode)
    SaveSetting REG_APP, REG_SEC_LAYOUT, REG_KEY_SLOT, CStr(mAnchorSlot)
End Sub

Private Sub btnCurrYear_Click()
    SetBaseYear Year(Date)   ' ���ɰ����� �Բ� ����ȭ
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

' ========= �ʱ�ȭ =========
Private Sub UserForm_Initialize()
    SetupColors
    
    LoadAllHolidaysIntoGlobalSet

    mYear = Year(Date)
    txtBaseYear.text = CStr(mYear)
    
    ' ������ ����
    Set HolidayByDate = CreateObject("Scripting.Dictionary")
    HolidayByDate.CompareMode = vbTextCompare
    Set HolidayYearLoaded = CreateObject("Scripting.Dictionary")
    
    Me.BackColor = vbWhite

    CreateAllMonthBlocks
    
    txtFromSel.BackColor = RGB(255, 244, 204)
    txtToSel.BackColor = RGB(255, 244, 204)
    
   
    ' === ���̾ƿ� ��ȣ�� �ε� ===
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
    
'    EnableMouseScroll Me, True, True, False          ' ����ó�� �� �� ��
'    WheelBridge.RegisterWheelSink Me   ' �� �߰�: �� ���� �� �����ڷ� ���
    
    SetBaseYear mYear, False
    
    InitTaskPanel
    
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
'    WheelBridge.UnregisterWheelSink Me ' �� �߰�: �� ���� �� ��� ����
End Sub

' 12ĭ ������ ǥ���� "�� ���� 1��" ��¥�� ä����
Private Sub BuildMonthMap(ByVal baseYear As Long, ByVal curMonth As Long, _
                          ByVal mode As eLayoutMode, ByRef slotDate() As Date, _
                          Optional ByVal anchorSlot As Long = 1)
    ReDim slotDate(1 To 12)

    Dim startD As Date
    Select Case mode
        Case lmNormal
            ' 1��~12�� (��� baseYear)
            startD = DateSerial(baseYear, 1, 1)

        Case lmCurrentFirst
            ' ����1 = baseYear�� ����� �� ���� 11����
            startD = DateSerial(baseYear, curMonth, 1)

        Case lmCurrentLast
            ' ����12 = baseYear�� ����� �� ���� 11����
            startD = DateAdd("m", -11, DateSerial(baseYear, curMonth, 1))

        Case lmCurrentAtSlot
            ' ����(anchorSlot) = baseYear�� �����
            If anchorSlot < 1 Or anchorSlot > 12 Then anchorSlot = 1
            ' ����1�� �ǵ��� ������� (-(anchorSlot-1))��ŭ ��� ���� ������
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
    ' �� ���� ���̶���Ʈ�� �ɼ� �ݿ��ؼ� �ٽ� ���� ĥ
    ApplySelectedRangeOverlay
End Sub

Private Sub RenderAllMonths()
    Dim baseYear As Long: baseYear = mYear
    
    Dim curM As Long: curM = Month(Date)  ' �ý��� ����� ���� ��Ŀ

    Dim mapDate() As Date
    
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot

    Dim i As Long
    For i = 1 To 12
        Dim y As Long, m As Long
        y = Year(mapDate(i))
        m = Month(mapDate(i))
        RenderMonthByBlock i, y, m    ' ���ο��� DrawMonthCells ȣ��
    Next
    
    EnsureMonthBGCreated
    SizeAllMonthBG
    UpdateCurrentMonthBG mapDate   ' �� ����� ��ϸ� ǥ��
    
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
    
    ' ���� ������ ���Ŀ� �ٽ� ���� ĥ�ϱ�
    ApplySelectedRangeOverlay
    
End Sub

' mBlock: 1..12 (ȭ����� ��ġ), y/m: ���� �޷��� ��/��
Private Sub RenderMonthByBlock(ByVal mBlock As Long, ByVal y As Long, ByVal m As Long)
    DrawMonthCells mBlock, y, m
End Sub

' mBlock : ȭ����� �� ��� �ε���(1..12)
' y, m   : ���� ǥ���� ��/��
Private Sub DrawMonthCells(ByVal mBlock As Long, ByVal y As Long, ByVal m As Long)
    Dim r As Long, c As Long, d As Long
    Dim firstDay As Date, lastDay As Long, startCol As Long
    Dim tb As MSForms.TextBox, dT As Date
    Dim tip As String, holName As String

    ' ==== ����(��� ��� ��Ÿ�� �Ҵ�) ====
    Dim CLR_BG As Long:            CLR_BG = RGB(255, 255, 255)
    Dim CLR_SUN_BG As Long:        CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:        CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_TODAY_BG As Long:      CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_HOLI_BG As Long:       CLR_HOLI_BG = RGB(255, 235, 235)

    Dim CLR_TEXT As Long:          CLR_TEXT = vbBlack
    Dim CLR_SUN_FG As Long:        CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:        CLR_SAT_FG = RGB(0, 90, 200)

    ' ==== Ÿ��Ʋ (yyyy��mm��) ====
    lblMonth(mBlock).Caption = MonthTitleOut(y, m)
    
    ' ==== Ÿ��Ʋ ���: ���� ¦/Ȧ�� ���� ====
    With lblMonth(mBlock)
        .BackStyle = fmBackStyleOpaque
        If (y Mod 2) = 0 Then
            .BackColor = mYearEvenBG   ' ¦����
        Else
            .BackColor = mYearOddBG    ' Ȧ����
        End If
        .ForeColor = mYearFG
    End With

    ' ==== ��� �� �ʱ�ȭ ====
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
                .Locked = True           ' ���� ����(Ŭ�� �̺�Ʈ�� ����)
                .BorderStyle = 0   ' �� �⺻�� ����
            End With
        Next
    Next

    ' ==== �� ��ġ ��� ====
    firstDay = DateSerial(y, m, 1)
    lastDay = Day(DateSerial(y, m + 1, 0))              ' ����
    startCol = Weekday(firstDay, FirstDOWParam())  ' Mon �����̸� vbMonday ����

    r = 1: c = startCol

    ' ==== ��¥ ä��� ====
    For d = 1 To lastDay
        Set tb = tbDay(mBlock, r, c)
        dT = DateSerial(y, m, d)

        ' �⺻ ��
        tb.text = CStr(d)
        tb.Tag = Format$(dT, "yyyy-mm-dd")
        tb.ControlTipText = FmtDateOut(dT) ' ��� ���� �ݿ�

        ' �ָ� �����/���� (��¥ ����)
        If IsSundayDate(dT) Then
            tb.ForeColor = CLR_SUN_FG
            tb.BackColor = CLR_SUN_BG
        ElseIf IsSaturdayDate(dT) Then
            tb.ForeColor = CLR_SAT_FG
            tb.BackColor = CLR_SAT_BG
        Else
            tb.ForeColor = CLR_TEXT
        End If

        ' ������(�̸� ��ȸ) �� ���/���� ����
        holName = GetHolidayNameIfAny(dT)
        If Len(holName) > 0 Then
            tb.BackColor = CLR_HOLI_BG
            tb.ForeColor = CLR_SUN_FG         ' ���������� �������� ���� �迭
            tb.ControlTipText = tb.ControlTipText & vbCrLf & "[������] " & holName
        End If

        ' ���� ����(���/����) - �����Ϻ��� �켱 ǥ���ϰ� ������ ���� ����
        If dT = Date Then
            tb.BackColor = CLR_TODAY_BG
            tb.Font.Bold = True
        End If

        ' ���� ���� �̵�
        c = c + 1
        If c > 7 Then c = 1: r = r + 1: If r > 6 Then Exit For
    Next

    ' ���ù��� ����ĥ �� ������ ȣ��
    On Error Resume Next
    PaintRange
    On Error GoTo 0
End Sub

Private Function GetHolidayNameIfAny(ByVal d As Date) As String
    On Error Resume Next
    If gHolidaySet Is Nothing Then Exit Function
    Dim k As Date: k = DateSerial(Year(d), Month(d), Day(d)) ' �ð� 0:00 ����ȭ
    If gHolidaySet.Exists(k) Then GetHolidayNameIfAny = CStr(gHolidaySet(k))
End Function

Public Sub HandleDayMouse(ByVal tb As MSForms.TextBox, ByVal Button As Integer)
    Dim s As String, d As Date
    s = Trim$(tb.Tag)
    If Len(s) = 0 Then Exit Sub                        ' ��ĭ ��
    If Not TryParseYMD(s, d) Then Exit Sub             ' ���
    
    If Button = 1 Then
        ' From
        mFrom = d: mHasFrom = True
    ElseIf Button = 2 Then
        ' To
        mTo = d: mHasTo = True
    Else
        Exit Sub
    End If

    ' From/To ����ȭ(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' ��� ǥ��
    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")

    PaintRange        ' ���� ���� �ٽ� ĥ�ϱ�
    UpdateRangeInfo   ' "Business Day / ���ϼ�" ����
    ApplySelectedRangeOverlay
    'RefreshTaskListAndOverlay
End Sub

' �ʿ� �� �� ���� ����
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
                .ZOrder 1      ' �ڷ� ������ (fmZOrderBack)
            End With
        End If
    Next
End Sub

' �� ��� 1..12�� �簢���� ����Ͽ� ��� �� ũ�� ����
Private Sub SizeAllMonthBG()
    Dim i As Long
    For i = 1 To 12
        SizeOneMonthBG i
    Next
End Sub

Private Sub SizeOneMonthBG(ByVal i As Long)
    On Error Resume Next
    If lblMonth(i) Is Nothing Then Exit Sub

    ' ����� �»��: Ÿ��Ʋ ��
    Dim leftEdge As Single, topEdge As Single
    leftEdge = lblMonth(i).Left
    topEdge = lblMonth(i).Top

    ' ����� ���ϴ�: 6�� 7�� TextBox (���� ����)
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
        .ZOrder 1   ' �ڷ�
    End With
End Sub

' �����(�ý��� ���� ����)�� ȭ�鿡 ������ �ش� ��ϸ� ���̰�
Private Sub UpdateCurrentMonthBG(ByRef mapDate() As Date)
    Dim i As Long, yT As Long, mT As Long
    yT = Year(Date): mT = Month(Date)

    ' ��� ����
    For i = 1 To 12
        If Not lblMonthBG(i) Is Nothing Then lblMonthBG(i).visible = False
    Next

    ' mapDate(i) = �� ������ "�� ���� 1��"
    For i = 1 To 12
        If Year(mapDate(i)) = yT And Month(mapDate(i)) = mT Then
            If Not lblMonthBG(i) Is Nothing Then
                lblMonthBG(i).visible = True
                lblMonthBG(i).ZOrder 1   ' �����ϰ� �ڷ�
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
    Call RefreshTaskOverlayIfOn   ' �� �߰�
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' �� �߰�
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentFirst_Click()
    mLayoutMode = lmCurrentFirst
    EnableAnchorSlotUI False
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' �� �߰�
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' �� �߰�
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentLast_Click()
    mLayoutMode = lmCurrentLast
    EnableAnchorSlotUI False
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' �� �߰�
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' �� �߰�
    RefreshTaskListAndOverlay
End Sub

Private Sub optCurrentAtSlot_Click()
    mLayoutMode = lmCurrentAtSlot
    EnableAnchorSlotUI True
    RenderAllMonths
    PaintRange
    Call RefreshTaskOverlayIfOn   ' �� �߰�
    SaveLayoutPrefs
    'LoadTasksForVisibleRangeAndOverlay   ' �� �߰�
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
            Call RefreshTaskOverlayIfOn   ' �� �߰�
            PaintRange
        End If
        SaveLayoutPrefs
        'LoadTasksForVisibleRangeAndOverlay   ' �� �߰�
        RefreshTaskListAndOverlay
    End If
End Sub

Private Sub RefreshTaskOverlayIfOn()
    ' ���� �ڿ� ������ �ֽ� �������� ������, On�̸� ����/Off�� ����
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

' ���� �ʱ�ȭ�� �ʿ��� �� ���(�ɼ�)
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


' ========= ���� ���� =========
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

        ' === �� ����� ��� ��(���� �����ؼ� �ڷ� ��ġ��) ===
        Dim bgW As Single, bgH As Single
        bgW = 7 * (CELL_W + CELL_GAP)                  ' �� ��ü ��
        bgH = TITLE_H + WEEK_H + 6 * (CELL_H + CELL_GAP) - CELL_GAP  ' ������ ���� ����
        Set lblMonthBG(m) = AddBGLabel("lblMonthBG" & m, originX - 4, originY - 4, bgW + 8, bgH + 8)

        ' === �� Ÿ��Ʋ ===
        Set lblMonth(m) = AddLabel("lblMonth" & m, originX, originY, 7 * (CELL_W + CELL_GAP) - 1, TITLE_H - 2, CStr(m) & "��", True)
        lblMonth(m).Font.name = "Segoe UI Semibold"
        lblMonth(m).Font.Size = FONT_SIZE + 2
        lblMonth(m).Font.Bold = True
        lblMonth(m).ForeColor = mClrMonthTitle

        ' === ���� ===
        Dim cidx As Long
        For cidx = 1 To 7
            Set lblWeek(m, cidx) = AddLabel("lblWeek" & m & "_" & cidx, _
                            originX + (cidx - 1) * (CELL_W + CELL_GAP), originY + TITLE_H, _
                            CELL_W, WEEK_H - 1, WeekNameByPos(cidx), True)
            lblWeek(m, cidx).BorderStyle = 0
            
            ' ���ϻ�(��/�丸 ����) - ���� ���� ����
            Dim realDow As Long: realDow = DowByPos(cidx)  ' 1=Sun .. 7=Sat
            If realDow = vbSunday Then
                lblWeek(m, cidx).ForeColor = mClrSunFg
            ElseIf realDow = vbSaturday Then
                lblWeek(m, cidx).ForeColor = mClrSatFg
            Else
                lblWeek(m, cidx).ForeColor = mClrWeekFg
            End If
        Next
        
        ' === ����(TextBox) 6��7 ===
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
    
    ' Control�� Event Hook
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
        .visible = False       ' �⺻�� ���� �� ������� ���̰�
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
        .Locked = True            ' ���� ����(Ŭ���� ����)
        .Enabled = True           ' ���콺 �̺�Ʈ�� ���� Enabled ����
    End With
    Set AddDayTextBox = t
End Function

' ========= ������ =========
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

                ' �⺻ ����
                ApplyBaseStyle tb, c

                ' ���� ������ ���� �������� ����
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

    ' �� �׻� �������� ������ (���� ������ ������ ����Ǿ� ����)
    RefreshTaskOverlayIfOn
End Sub


' ���� ���¿����� From~To ������ ���
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


' ��(colIndex) ������ �⺻ ��Ÿ�� ���� ��,
' ��ȿ ��¥(=Tag�� yyyy-mm-dd)�� ���� �ָ�/������/���� ���� ����
Private Sub ApplyBaseStyle(ByVal tb As MSForms.TextBox, ByVal colIndex As Long)
    ' �⺻ �ȷ�Ʈ
    Dim CLR_BG As Long:       CLR_BG = vbWhite
    Dim CLR_TEXT As Long:     CLR_TEXT = vbBlack
    Dim CLR_SUN_BG As Long:   CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:   CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_TODAY_BG As Long: CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_HOLI_BG As Long:  CLR_HOLI_BG = RGB(255, 235, 235)
    Dim CLR_SUN_FG As Long:   CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:   CLR_SAT_FG = RGB(0, 90, 200)

    ' 1) �⺻ �ʱ�ȭ
    tb.Font.Bold = False
    tb.ForeColor = CLR_TEXT
    tb.BackColor = CLR_BG

    ' 2) ��¥ ���� ĭ(=Tag ����/��ȿ)�� ���⼭ �� �� �׻� �� ��� ����
    Dim s As String: s = Trim$(tb.Tag)
    If Len(s) = 0 Then Exit Sub

    Dim d As Date
    If Not TryParseYMD(s, d) Then Exit Sub

    ' 3) ��ȿ ��¥�� ���� �ָ�/���� ��
    If IsSundayDate(d) Then
        tb.ForeColor = CLR_SUN_FG
        tb.BackColor = CLR_SUN_BG
    ElseIf IsSaturdayDate(d) Then
        tb.ForeColor = CLR_SAT_FG
        tb.BackColor = CLR_SAT_BG
    End If

    ' 4) ������ ����(�̸� ������)
    Dim nm As String: nm = GetHolidayNameIfAny(d)
    If Len(nm) > 0 Then
        tb.BackColor = CLR_HOLI_BG
        tb.ForeColor = CLR_SUN_FG
    End If

    ' 5) ���� ����
    If d = Date Then
        tb.BackColor = CLR_TODAY_BG
        tb.Font.Bold = True
    End If
End Sub



' ========= UI �̺�Ʈ =========
Private Sub txtBaseYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH
    Dim y As Long: y = CLng(Trim$(txtBaseYear.text))
    If y < 1901 Or y > 9998 Then GoTo EH
    SetBaseYear y                     ' �߾� �Լ��� �Ͽ�ȭ
    Exit Sub
EH:
    MsgBox "���س⵵�� ��Ȯ�� �Է��ϼ���. ��) 2025", vbExclamation
    txtBaseYear.text = CStr(mYear)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ========= ������ �ε�/���� =========
Private Sub EnsureHolidayYearLoaded(ByVal y As Long)
    Dim ky As String: ky = CStr(y)
    If HolidayYearLoaded.Exists(ky) Then Exit Sub

    Dim raw As String: raw = ReadHolidaysRaw(y) ' "yyyy-mm-dd|���ϸ�" �ٵ�
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

' ����Ŭ��: �̼����� ������ ä��� �ٷ� ��ȯ
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
        ' �� �� ������ ������ �״�� �ΰ� ��� ��ȯ
    End If

    ' From<=To ����
    Dim t As Date
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then t = mFrom: mFrom = mTo: mTo = t
    End If

    txtFromSel.text = IIf(mHasFrom, FmtDateOut(mFrom), "")
    txtToSel.text = IIf(mHasTo, FmtDateOut(mTo), "")

    PaintRange
    PushRangeNow
End Sub

' ��ȯ ����(��ư/����Ŭ�� ����)
Private Sub PushRangeNow()
    If Not mHasFrom And Not mHasTo Then
        MsgBox "���õ� ��¥�� �����ϴ�.", vbExclamation: Exit Sub
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
'                    MsgBox "��� ������ �����ϴ�. ���� �����ϰų� SetTargetRange�� �����ϼ���.", vbExclamation
'                    Exit Sub
'                End If
'            Else
                Set rng = Selection
                On Error GoTo 0
                If rng Is Nothing Then
                    MsgBox "��� ������ �����ϴ�. ���� ���� �ϼ���.", vbExclamation
                    Exit Sub
                End If
'            End If
            
            If chkFromOnly Then
                If Not mHasFrom Then
                    MsgBox "���õ� ��¥�� �����ϴ�.", vbExclamation: Exit Sub
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
            ' �⺻: Selection�� ��
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

' ���� ���� ������ �Ʒ��� 1�� �̵�(ũ�� ����)
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


' ����(��ȯ)
Private Sub btnApplyRange_Click()
    PushRangeNow
End Sub

' Clear: From/To �ʱ�ȭ + ĥ�� �� �ǵ�����
Private Sub btnClear_Click()
    ClearSelectionUI
End Sub

' ==== ������ ���� (������Ʈ�� �ε��� �� gHolidaySet ���) ====
Private Function IsHoliday(ByVal d As Date) As Boolean
    On Error Resume Next
    If gHolidaySet Is Nothing Then Exit Function
    ' �������� ����ȭ(�ð��� 0:00 �� �ƴϾ ����)
    IsHoliday = gHolidaySet.Exists(CDate(Int(CDbl(d))))
End Function

' ==== Business Day ����(�� �� ����) ====
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

' ==== ǥ�� ������Ʈ: "Business Day / ���ϼ�" ====
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

' ��/���� ���� ���ڿ� ��ȯ
Private Function NzCStr(ByVal v As Variant) As String
    If IsError(v) Or isNull(v) Or IsEmpty(v) Then NzCStr = "" Else NzCStr = CStr(v)
End Function

Private Sub btnExport_Click()
    ExportCalendarCurrentLayout
End Sub

' ���� ���̾ƿ� �״�� �� ��Ʈ�� ���
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
    
    ' �ȷ�Ʈ(��Ÿ�� ����)
    Dim CLR_TEXT As Long:     CLR_TEXT = vbBlack
    Dim CLR_BG As Long:       CLR_BG = vbWhite
    Dim CLR_GRID As Long:     CLR_GRID = RGB(210, 210, 210)
    Dim CLR_SUN_FG As Long:   CLR_SUN_FG = RGB(200, 0, 0)
    Dim CLR_SAT_FG As Long:   CLR_SAT_FG = RGB(0, 90, 200)
    Dim CLR_SUN_BG As Long:   CLR_SUN_BG = RGB(255, 248, 248)
    Dim CLR_SAT_BG As Long:   CLR_SAT_BG = RGB(245, 248, 255)
    Dim CLR_HOLI_BG As Long:  CLR_HOLI_BG = RGB(255, 235, 235)
    Dim CLR_TODAY_BG As Long: CLR_TODAY_BG = RGB(230, 255, 230)
    Dim CLR_SEL_BG As Long:   CLR_SEL_BG = RGB(33, 92, 152) ' From~To ����  CLR_SEL_BG = RGB(255, 244, 204) ' From~To ����
    Dim CLR_SEL_FG As Long:   CLR_SEL_FG = RGB(255, 255, 255) ' From~To ����
    
    Dim YearOddBG As Long:    YearOddBG = RGB(255, 248, 230)
    Dim YearEvenBG As Long:   YearEvenBG = RGB(236, 244, 255)
    Dim YearFG As Long:       YearFG = RGB(30, 30, 30)
    
    ' ��ġ ��Ʈ����
    Const COL0 As Long = 2         ' ù ��� �»�� �÷�
    Const ROW0 As Long = 4         ' ù ��� �»�� �� (1~3���� ���/��� ��)
    Const COLS_PER_MONTH As Long = 7  ' ���� 7ĭ
    Const ROWS_TITLE As Long = 1
    Const ROWS_WEEK As Long = 1
    Const ROWS_DAYS As Long = 6
    Const ROWS_PER_MONTH As Long = ROWS_TITLE + ROWS_WEEK + ROWS_DAYS
    Const GAP_COLS As Long = 2
    Const GAP_ROWS As Long = 1
    
    'Dim weekHdr: weekHdr = Split("��,��,ȭ,��,��,��,��", ",")
    
    ' ���ؿ���/����� & ��(12ĭ)
    Dim baseYear As Long: baseYear = CLng(Val(txtBaseYear.text))
    Dim curM As Long: curM = Month(Date)
    Dim mapDate() As Date
    
    Application.ScreenUpdating = False
    
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot   ' (��Ŀ ���� ��� ����)
    
    BuildDayBoxMapFromGrid    ' �� UI�� ���̴� DayBox ���� �ֽ�ȭ
    
    Dim have As Boolean, sRange As Date, eRange As Date
    Call GetSelectedRange(have, sRange, eRange)
    
    ' ��� ���(���̾ƿ�/����)
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
            ws.Cells(3, 2).Value = "���ù���: " & FmtDateOut(sRange) & " ~ " & FmtDateOut(eRange) _
                                   & "  (" & CStr(biz) & " / " & CStr(totalDays) & ")"
        End If
        ws.Cells(3, 2).Font.Bold = True
    Else
        If mLang = LangE Then
            ws.Cells(3, 2).Value = "Range: (Nothing)"
        Else
            ws.Cells(3, 2).Value = "���ù���: (����)"
        End If
        
    End If
    
    ' ��Ʈ �⺻
    With ws.Cells
        .Font.name = "Segoe UI"
        .Font.Size = 8
    End With
    
    ' === 12���� ���� ===
    Dim i As Long, y As Long, m As Long
    For i = 1 To 12
        y = Year(mapDate(i)): m = Month(mapDate(i))
        
        Dim rowBlock As Long, colBlock As Long
        rowBlock = (i - 1) \ GRID_COLS
        colBlock = (i - 1) Mod GRID_COLS
        
        Dim c0 As Long, r0 As Long
        c0 = COL0 + colBlock * (COLS_PER_MONTH + GAP_COLS)
        r0 = ROW0 + rowBlock * (ROWS_PER_MONTH + GAP_ROWS)
        
        ' --- Ÿ��Ʋ ---
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
        
         ' --- ���� ��� (�� ����/ǥ�� ��Ÿ�� �ݿ�) ---
        Dim j As Long, realDow As Long
        For j = 1 To 7
            With ws.Cells(r0 + ROWS_TITLE, c0 + (j - 1))
                .Value = WeekNameByPos(j)                 ' �� ���� �ؽ�Ʈ(Short/Full, K/E �ݿ�)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = vbWhite
        
                realDow = DowByPos(j)                      ' 1=Sun .. 7=Sat(�� ���ۿ� ���� ����)
                If realDow = vbSunday Then
                    .Font.Color = CLR_SUN_FG               ' �Ͽ��� ����
                ElseIf realDow = vbSaturday Then
                    .Font.Color = CLR_SAT_FG               ' ����� �Ķ�
                Else
                    .Font.Color = CLR_TEXT                 ' ���� �⺻��
                End If
        
                .RowHeight = 14
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = CLR_GRID
                End With
            End With
        Next
        
        ' --- ��¥ ä��� ---
        Dim firstD As Date: firstD = DateSerial(y, m, 1)
        Dim lastDay As Long: lastDay = Day(DateSerial(y, m + 1, 0))
        
        Dim startCol As Long: startCol = Weekday(firstD, FirstDOWParam())
        
         ' --- ��¥ ä��� (������ �޸� ����) ---
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
        
                ' �ָ� ��
                If IsSundayDate(dd) Then
                    .Font.Color = CLR_SUN_FG
                    .Interior.Color = CLR_SUN_BG
                ElseIf IsSaturdayDate(dd) Then
                    .Font.Color = CLR_SAT_FG
                    .Interior.Color = CLR_SAT_BG
                End If
                
                If Not mTaskOverlay Then
                    ' ������ �޸� �� note�� ���� (���� .AddComment ����)
                    If IsHolidayWS(dd) Then
                        .Interior.Color = CLR_HOLI_BG
                        .Font.Color = CLR_SUN_FG
                        Dim holName As String
                        holName = GetHolidayNameIfAny(dd)
                        If Len(holName) > 0 Then
                            Dim holPrefix As String
                            holPrefix = IIf(mLang = LangE, "[Holiday] ", "[������] ")
                            If Len(note) > 0 Then note = note & vbCrLf
                            note = note & holPrefix & holName
                        End If
                    End If
                End If
        
                ' ���� ����
                If dd = Date Then
                    .Interior.Color = CLR_TODAY_BG
                    .Font.Bold = True
                End If
        
                ' ���� ���� ��������
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
                
                ' === �� �������� ����: UI DayBox�� ���(���/�۲�/�׵θ�/����)�� �״�� ��Ʈ�� �ݿ� ===
                Dim tbUI As MSForms.TextBox
                Set tbUI = FindDayBoxByDate(dd)   ' UI �� �ش� ��¥ ĭ
                If Not tbUI Is Nothing Then
                    On Error Resume Next
                    ' 1) ���/�۲�(���ù���/�������̰� �̹� �ݿ��� ���� ����)
                    .Interior.Color = tbUI.BackColor
                    .Font.Bold = tbUI.Font.Bold
                    .Font.Color = tbUI.ForeColor
                
                    ' 2) �������� �׵θ��� �ݿ�(���� ��)
                    If tbUI.BorderStyle <> 0 Then
                        With .Borders
                            .LineStyle = xlContinuous
                            .Color = tbUI.BorderColor
                            .Weight = xlThin
                        End With
                    End If
                    
                    If mTaskOverlay Then
                        ' 3) �������� ������ �޸� ����
                        If Len(tbUI.ControlTipText) > 0 Then
                            If Len(tbUI.ControlTipText) = 10 Then
                            Else
                                If Len(note) > 0 Then note = note & vbCrLf
                                'note = note & IIf(mLang = LangE, "[Tasks] ", "[�۾�] ") & Mid(tbUI.ControlTipText, 11)
                                note = note & Mid(tbUI.ControlTipText, 11)
                            End If
                        End If
                    End If
                
'                    ' 3) �������� ������ �޸� ����
'                    If Len(tbUI.ControlTipText) > 0 Then
'                        If Len(note) > 0 Then note = note & vbCrLf
'                        note = note & IIf(mLang = LangE, "[Tasks] ", "[�۾�] ") & tbUI.ControlTipText
'                    End If
                    On Error GoTo 0
                End If
                
                ' === �� ���� �޸� ���(������+�������� �պ�) ===
                If Len(note) > 0 Then
                    On Error Resume Next
                    If Not .Comment Is Nothing Then .Comment.Delete
                    .AddComment note
                    .Comment.visible = False
                    With .Comment.Shape.TextFrame
                        .AutoSize = True
                        .Characters.Font.name = "���� ���"
                        .Characters.Font.Size = 9
                    End With
                    On Error GoTo 0
                End If
        
                ' �׵θ�
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = CLR_GRID
                End With
            End With
        
            ' ���� ĭ
            cc = cc + 1
            If cc > 7 Then cc = 1: rr = rr + 1: If rr >= ROWS_DAYS Then Exit For
        Next d
        
        ' ��ĭ(�� �� ����)�� ��/���� ��� ä �׵θ��� ��� ����(������)
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
        
        ' Į�� �� ����ȭ(�� ����� 7��)
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

' ���� gHolidaySet(Dictionary: Key=Date, Val=Name)�� ���
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
'' frmYearCalendar ���� �߰�
'Public Sub OnMouseWheel(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)
'    ' dir: +1(��), -1(�Ʒ�)�� ���ɴϴ�.
'    If dir = 0 Then Exit Sub
'
'    Dim stepVal As Long: stepVal = 1
'    ' ���� RouteKey(PageUp/Down) ��Ģ�� ������ ����:
'    '  Shift+Ctrl: 15,  Ctrl: 5,  Shift: 10,  �⺻: 1
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
'    ' spnYear_Change ���� �̹� RenderAllMonths/PaintRange ȣ���ϹǷ� �߰� ȣ�� ���ʿ�
'End Sub

Private Sub SetFromDate(d As Date)
    mFrom = d
    mHasFrom = True

    ' From/To ����ȭ(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' ��� ǥ��
    txtFromSel.text = IIf(mHasFrom, Format$(mFrom, "yyyy-mm-dd"), "")

    PaintRange        ' ���� ���� �ٽ� ĥ�ϱ�
    UpdateRangeInfo   ' "Business Day / ���ϼ�" ����
    ApplySelectedRangeOverlay
End Sub

Public Sub SetToDate(d As Date)
    mTo = d
    mHasTo = True

    ' From/To ����ȭ(From<=To)
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim tmp As Date
            tmp = mFrom
            mFrom = mTo
            mTo = tmp
        End If
    End If

    ' ��� ǥ��
    txtToSel.text = IIf(mHasTo, Format$(mTo, "yyyy-mm-dd"), "")

    PaintRange        ' ���� ���� �ٽ� ĥ�ϱ�
    UpdateRangeInfo   ' "Business Day / ���ϼ�" ����
    ApplySelectedRangeOverlay
End Sub


' frmYearCalendar �� ���� ��ġ(��� ����) �߰�
Public Sub HandleMonthLabelMouse(ByVal L As MSForms.Label, ByVal Button As Integer)
    Dim idx As Long
    Dim baseYear As Long, curM As Long
    Dim mapDate() As Date
    Dim y As Long, m As Long
    Dim d As Date

    ' ���̺� �̸�: "lblMonth" & �ε���(1..12)
    idx = CLng(Val(Mid$(L.name, 9)))
    If idx < 1 Or idx > 12 Then Exit Sub

    ' ���� ȭ���� 12ĭ�� ����Ű�� ���� (��/��) ����
    baseYear = CLng(Val(txtBaseYear.text))
    curM = Month(Date)
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot

    y = Year(mapDate(idx))
    m = Month(mapDate(idx))

    Select Case Button
        Case 1 ' Left �� From = �� ���� 1��
            d = DateSerial(y, m, 1)
            mFrom = d: mHasFrom = True
        Case 2 ' Right �� To = �� ���� ����
            d = DateSerial(y, m + 1, 0)
            mTo = d: mHasTo = True
        Case Else
            Exit Sub
    End Select

    ' From �� To ����
    If mHasFrom And mHasTo Then
        If mTo < mFrom Then
            Dim t As Date: t = mFrom: mFrom = mTo: mTo = t
        End If
    End If

    ' UI �ݿ� �� ĥ�ϱ�
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

    ' ��¥ ����(����; �ʿ�� �и� ����)
    mFmtDate = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_DATE, _
                 IIf(mLang = LangE, DEF_FMT_DATE_E, DEF_FMT_DATE_K))

    ' ===== �� Ÿ��Ʋ ����(�� �и� ����) =====
    ' �� Ÿ��Ʋ ���� �ε�
    Dim rawK As String, rawE As String, legacy As String
    rawK = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE_K, DEF_FMT_TITLE_K)
    rawE = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE_E, DEF_FMT_TITLE_E)
    
    ' (�ɼ�) ���Ž� Ű�� ������ �ʱ� �̰�
    legacy = GetSetting(REG_APP, REG_SEC_FMT, REG_KEY_FMT_TITLE, "")
    If legacy <> "" Then
        If rawK = DEF_FMT_TITLE_K Then rawK = legacy
        If rawE = DEF_FMT_TITLE_E Then rawE = legacy
    End If
    
    ' ������ ���� ���
    mFmtMonthTitleK = SanitizeMonthTitleFmt(rawK, False)
    mFmtMonthTitleE = SanitizeMonthTitleFmt(rawE, True)
    
    ' 3) ���� �� �°� ����
    mFmtMonthTitle = IIf(mLang = LangE, mFmtMonthTitleE, mFmtMonthTitleK)

    ' --- WeekStart / WeekNameStyle (���� �亯���� �߰��� �κ� ����) ---
    Dim ws As String, wn As String
    ws = GetSetting(REG_APP, REG_SEC_I18N, REG_KEY_WEEK_START, "Sun")
    wn = GetSetting(REG_APP, REG_SEC_I18N, REG_KEY_WEEKNAME_STYLE, "Short")
    mWeekStart = IIf(UCase$(ws) = "MON", WeekMon, WeekSun)
    mWeekNameStyle = IIf(UCase$(wn) = "FULL", WkFull, WkShort)
End Sub

Private Function SanitizeMonthTitleFmt(ByVal s As String, ByVal isEnglish As Boolean) As String
    Dim t As String: t = Trim$(s)
    If t = "" Then GoTo def
    ' ���� ��ū(d/D)�� �� ���� �� Ÿ��Ʋ������ ������ �� �⺻������
    If InStr(1, LCase$(t), "d", vbTextCompare) > 0 Then GoTo def
    SanitizeMonthTitleFmt = t
    Exit Function
def:
    SanitizeMonthTitleFmt = IIf(isEnglish, DEF_FMT_TITLE_E, DEF_FMT_TITLE_K)
End Function

Private Sub SaveLangToRegistry()
    SaveSetting REG_APP, REG_SEC_I18N, REG_KEY_LANG, IIf(mLang = LangE, "E", "K")
End Sub

' �� ���� ��Ī
Private Function WeekNameByIndex(ByVal idx As Long) As String
    ' idx: 1=Sun ~ 7=Sat
    If mLang = LangE Then
        WeekNameByIndex = Split("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")(idx - 1)
    Else
        WeekNameByIndex = Split("��,��,ȭ,��,��,��,��", ",")(idx - 1)
    End If
End Function

' �� Ÿ��Ʋ ���ڿ�
Private Function MonthTitleOut(ByVal y As Long, ByVal m As Long) As String
    MonthTitleOut = Format$(DateSerial(y, m, 1), mFmtMonthTitle)
End Function

' ��¥ ��� ���ڿ� (�ؽ�Ʈ�ڽ�/��ȯ��)
Private Function FmtDateOut(ByVal d As Date) As String
    FmtDateOut = Format$(d, mFmtDate)
End Function

' Excel ���� ��¥ ����(��+ǥ������) ? ��� �� ��� �������� �̰� ȣ��
Private Sub ApplyDateToCell(ByVal tgt As Range, ByVal d As Date)
    If tgt Is Nothing Then Exit Sub
    tgt.Value = d
    On Error Resume Next
    tgt.NumberFormatLocal = mFmtDate   ' ��κ� Excel�� ���� ��ū. ���� ���� ���
    On Error GoTo 0
End Sub

' UI ���� �ؽ�Ʈ(��ư/üũ�ڽ� ĸ�� ��) ��/�� �ݿ�
Private Sub ApplyStaticUIStrings()
    On Error Resume Next

    ' === ���� K/E ��� üũ�ڽ� ===
    chkKE.Caption = "K/E"

    ' ��ư/üũ�ڽ�/�ɼ� �� (�ʿ��� ������ ����)
    If mLang = LangE Then
        btnCurrYear.Caption = "ThisYear"
        btnApplyRange.Caption = "Apply"
        btnClear.Caption = "Clear"
        btnClose.Caption = "Close"
        btnExport.Caption = "Export"
        btnConfig.Caption = "Set"

        optNormal.Caption = "1��12"
        optCurrentFirst.Caption = "Current First"
        optCurrentLast.Caption = "Current Last"
        optCurrentAtSlot.Caption = "Current @Slot"

        chkFromOnly.Caption = "From Only"
        chkShowKeep.Caption = "Keep form open"
        chkExcludeNonBiz.Caption = "Exclude Sat/Sun/Holidays in range color"
        chkAutoNextRow.Caption = "Go��"

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
        btnCurrYear.Caption = "����"
        btnApplyRange.Caption = "����(��ȯ)"
        btnClear.Caption = "Clear"
        btnClose.Caption = "�ݱ�"
        btnExport.Caption = "���"
        btnConfig.Caption = "����"

        optNormal.Caption = "1����12��"
        optCurrentFirst.Caption = "������� ���ۿ�"
        optCurrentLast.Caption = "������� ����"
        optCurrentAtSlot.Caption = "����� @Slot"

        chkFromOnly.Caption = "From Only"
        chkShowKeep.Caption = "â ����"
        chkExcludeNonBiz.Caption = "����ǥ�ÿ� ���� ����"
        chkAutoNextRow.Caption = "�Ʒ��̵�"

        lblFromTo.Caption = "From~To"
        lblBaseYear.Caption = "����"
        lblSlot.Caption = "����"
        
        btnTaskLoad.Caption = "�ҷ�����"
        btnTaskSaveAppend.Caption = "�߰� ����"
        
        btnTaskGet.Caption = "��Ʈ ���ÿ������� �ҷ�����"
        btnTaskExport.Caption = "��Ʈ�� ���"
        chkTaskOverlay.Caption = "Task ������ Calendar�� ǥ��"
        chkTaskLinkSel.Caption = "Task ���ý� From~To�� �ݿ�"
        chkTaskShowAll.Caption = "��ü ����" & vbLf & "(������ �� Calendar ���� Task�� ǥ��)"
        btnRemoveTaskYear.Caption = "���� ���� ��� ����"
        btnTaskDeleteSelected.Caption = "���� �׸� ����"
        
        btnTaskAdd.Caption = "�߰�"
        btnTaskUpdate.Caption = "����"
        btnTaskDelete.Caption = "����"
        btnTaskSort.Caption = "����"
        
    End If
End Sub

' ���� ���/Ÿ��Ʋ �� ������
Private Sub ReapplyLanguageOnCalendar()
    Dim i As Long, j As Long
    ' �� Ÿ��Ʋ
    For i = 1 To 12
        Dim y As Long, m As Long, baseYear As Long, curM As Long, mapDate() As Date
        baseYear = CLng(Val(txtBaseYear.text))
        curM = Month(Date)
        BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot
        y = Year(mapDate(i)): m = Month(mapDate(i))
        If Not lblMonth(i) Is Nothing Then lblMonth(i).Caption = MonthTitleOut(y, m)
        ' ���� ���
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
    mLang = IIf(chkKE.Value, LangE, LangK)   ' üũ=����, ����=�ѱ�
    SaveLangToRegistry
    ' ��� �ٲ�� ���� �⺻���� �� �� ����ȭ(����ڰ� �����ص� ���� ������ �״�� ����)
    If mFmtDate = "" Then mFmtDate = IIf(mLang = LangE, DEF_FMT_DATE_E, DEF_FMT_DATE_K)
    If mFmtMonthTitle = "" Then mFmtMonthTitle = IIf(mLang = LangE, DEF_FMT_TITLE_E, DEF_FMT_TITLE_K)

    ApplyStaticUIStrings
    ' �� Ÿ��Ʋ/���� ������ + �� ���� ����
    ReapplyLanguageOnCalendar
    PaintRange
    
    SaveLangToRegistry
    LoadI18NAndFormats        ' �� ��� �ٲٸ� �� Ÿ��Ʋ ���˵� �� ������ ��ε�
    ReapplyLanguageOnCalendar ' �� �� Ÿ��Ʋ/���� ĸ�� �ٽ� �ݿ�
    
End Sub

Private Sub btnConfig_Click()
    frmDateFormatConfig.Show vbModal
    ' ����ڰ� �����ߴٸ� ������Ʈ������ �ٽ� �а� �ݿ�
    LoadI18NAndFormats
    ApplyStaticUIStrings
    RenderAllMonths          ' �� �� ���� �ٲ�� ��ġ�� �޶����Ƿ� �緻��
    ReapplyLanguageOnCalendar
    ' ���� �ؽ�Ʈ�� �� �������� �ٽ� ǥ��
    If mHasFrom Then txtFromSel.text = FmtDateOut(mFrom)
    If mHasTo Then txtToSel.text = FmtDateOut(mTo)
    PaintRange
End Sub

' �� ���� ������ VbDayOfWeek�� ȯ��
Private Function FirstDOWParam() As VbDayOfWeek
    FirstDOWParam = IIf(mWeekStart = WeekMon, vbMonday, vbSunday)
End Function

' ȭ����� ��ġ(1..7) �� ���� ���Ϲ�ȣ(1=Sun .. 7=Sat)
Private Function DowByPos(ByVal pos As Long) As Long
    ' WeekSun: pos=1��Sun,2��Mon,...7��Sat
    ' WeekMon: pos=1��Mon(2), ..., 6��Sat(7), 7��Sun(1)
    If mWeekStart = WeekSun Then
        DowByPos = pos
    Else
        DowByPos = ((pos Mod 7) + 1)   ' 1��2, 2��3, ..., 6��7, 7��1
    End If
End Function

' ���ϸ�(���/��Ÿ�� �ݿ�)
Private Function WeekNameByDow(ByVal dow As Long) As String
    If mLang = LangE Then
        If mWeekNameStyle = WkFull Then
            WeekNameByDow = Split("Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", ",")(dow - 1)
        Else
            WeekNameByDow = Split("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")(dow - 1)
        End If
    Else
        ' �ѱ��� ���� ��Ī ����(���Ͻø� '�Ͽ���,������,...'�� Full �߰� ����)
        If mWeekNameStyle = WkFull Then
            WeekNameByDow = Split("�Ͽ���,������,ȭ����,������,�����,�ݿ���,�����", ",")(dow - 1)
        Else
            WeekNameByDow = Split("��,��,ȭ,��,��,��,��", ",")(dow - 1)
        End If
    End If
End Function

' ȭ�� ��ġ(1..7) �� ���ϸ�
Private Function WeekNameByPos(ByVal pos As Long) As String
    WeekNameByPos = WeekNameByDow(DowByPos(pos))
End Function

' ��¥ ��� �ָ� ����(�� ���۰� ����)
Private Function IsSundayDate(ByVal d As Date) As Boolean
    IsSundayDate = (Weekday(d, vbSunday) = vbSunday)
End Function
Private Function IsSaturdayDate(ByVal d As Date) As Boolean
    IsSaturdayDate = (Weekday(d, vbSunday) = vbSaturday)
End Function

Private Sub btnHelp_Click()
    ' �𵨸����� �ΰ� �޷°� ���ÿ� �����Ϸ��� vbModeless
    frmYearCalHelp.Show vbModeless
End Sub

' ���� ������ �� �������� ó��: mYear, txtBaseYear, spnYear.Value ����ȭ + (�ɼ�)����
Private Sub SetBaseYear(ByVal y As Long, Optional ByVal doRender As Boolean = True)
    Dim newY As Long
    newY = y
    If newY < 1901 Then newY = 1901
    If newY > 9998 Then newY = 9998

    If mSyncingYear Then Exit Sub
    mSyncingYear = True
    On Error Resume Next

    ' ���� ���� ����
    If spnYear.Min <> 1901 Then spnYear.Min = 1901
    If spnYear.Max <> 9998 Then spnYear.Max = 9998

    ' ���� ���� ����
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
' === �ʱ�ȭ ������ ȣ�� (���� Initialize ���κп� �̾ �־��ּ���) ===
Private Sub InitTaskPanel()
    ' "NOHOOK" �±װ� Panel ��Ʈ�ѿ� �����Ǿ� �־�� ��
    ' ListBox �÷� ����
    With Me.lstTask
        .ColumnCount = 3
        .ColumnHeads = False
        ' 3��: Name | From | To
        '.ColumnWidths = "100;20;20"
        .BoundColumn = 0
        .MultiSelect = fmMultiSelectExtended   ' �� ���� ���� ���
    End With

    ' �� �ʺ�/����� ���� (������ �������� �����ϼ���)
    mBaseWidth = 700 ' ���� ������ �⺻�� (�ʿ�� �� ���� Width�� ��ü)
    mTaskWidth = 975 ' ��ģ ������ ��
    
    ' === �� ������Ʈ������ ���� ���� �ε� ===
    mTaskVisible = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_VISIBLE, True)
    mTaskOverlay = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_OVERLAY, False)
    mTaskLinkSel = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_TASK_LINKSEL, True)
    
    EnsureTaskPanelVisible mTaskVisible

    ' üũ�ڽ� ���� �ʱ�ȭ(������Ʈ�� ��� �ҷ��͵� ��)
    Me.chkTaskOverlay.Value = mTaskOverlay
    Me.chkTaskLinkSel.Value = mTaskLinkSel

    ' Ķ���� DayBox ���� �غ�
    BuildDayBoxMapFromGrid
    
    ' �� �������� ���س⵵ Task �ڵ� �ε� & �������� �ݿ�
    LoadTasksForVisibleRangeAndOverlay
    
    ' ... ���� �ڵ� ...
    ' === ��ü ���� �ʱ� ����(������Ʈ�� ����) ===
    Me.chkTaskShowAll.Value = GetRegBool(REG_SEC_TASKPANEL, REG_KEY_SHOWALL, False)
    
    ' === ó�� �ε� ===
    LoadCategoryList
    RefreshTaskListAndOverlay   ' (�Ʒ� �Լ�)
    
End Sub

Private Sub EnsureTaskPanelVisible(ByVal visible As Boolean)
    mTaskVisible = visible
    Me.Width = IIf(visible, mTaskWidth, mBaseWidth)
    Dim en As Boolean: en = visible
    ' �г� ��Ʈ�� ��Ŀ��/��/Ȱ�� ����
    PanelSetEnabled en
End Sub

Private Sub PanelSetEnabled(ByVal en As Boolean)
    Dim ctl As MSForms.control
    For Each ctl In Me.Controls
        If IsTaskPanelControl(ctl) Then
            On Error Resume Next
            ctl.Enabled = en
            ' �� ���� ����
            If HasProp(ctl, "TabStop") Then
                CallByName ctl, "TabStop", VbLet, en
            End If
            On Error GoTo 0
        End If
    Next
End Sub

Private Function IsTaskPanelControl(ByVal ctl As MSForms.control) As Boolean
    ' �����ν� Task Panel ���� ��Ʈ�� ���� Tag="NOHOOK" �ο�
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

'=== tbDay(12x6x7) ��ȸ�Ͽ� "yyyy-mm-dd" �� TextBox ���� ===
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
                        ' ���� ������ ĳ��
                        If OrigToolTip.Exists(key) Then
                            OrigToolTip(key) = tb.ControlTipText
                        Else
                            OrigToolTip.add key, tb.ControlTipText
                        End If
                        ' �⺻�� ���� ��Ģ: ���⼭�� BorderStyle ���� X
                    End If
                End If
            Next c
        Next r
    Next m
End Sub

'=== Tag�� "yyyy-mm-dd" �̸� True, ǥ��Ű ��ȯ ===
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

' Name�� "tbD_" �� �����ϸ� �ڿ��� ��¥�� �̾� "yyyy-mm-dd"�� ��ȯ
' ��� ��: tbD_20250107, tbD_2025-01-07, tbD_2025_1_7 ��
Private Function TryGetDateKeyFromName(ByVal nm As String, ByRef key As String) As Boolean
    On Error GoTo EH
    TryGetDateKeyFromName = False: key = ""
    If Left$(nm, 4) <> "tbD_" Then Exit Function

    Dim rest As String: rest = Mid$(nm, 5)
    ' 1) 8�ڸ� ����(yyyymmdd)
    If Len(rest) = 8 And IsNumeric(rest) Then
        key = Format$(DateSerial(CLng(Left$(rest, 4)), _
                                 CLng(Mid$(rest, 5, 2)), _
                                 CLng(Mid$(rest, 7, 2))), "yyyy-mm-dd")
        TryGetDateKeyFromName = True
        Exit Function
    End If
    ' 2) ������ ����: ���ڸ� ���� �� 3��ū(y,m,d)
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

'=== ��� ��ư ===
Private Sub btnTaskToggle_Click()
    Dim newVisible As Boolean
    newVisible = Not mTaskVisible
    EnsureTaskPanelVisible newVisible
    btnTaskToggle.Caption = IIf(mTaskVisible, "��", "��") ' ������ ���¸� �ݴ� ������(��), ���� ���¸� ���� ������(��)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_VISIBLE, mTaskVisible
End Sub


Private Function GetCurrentYearForTask() As Long
    ' ���� ���� "���ؿ���" ���� ������ �°� ���� ��ȯ�ϼ���.
    ' ����: txtYear ��� TextBox�� �ִٰ� ����
    On Error Resume Next
    Dim y As Long
    y = CLng(Me.txtBaseYear.Value)
    If y < 1900 Or y > 9999 Then y = Year(Date)
    GetCurrentYearForTask = y
End Function

' === Task Panel: [�߰� ����] ===
Private Sub btnTaskSaveAppend_Click()
    On Error GoTo EH

    Dim cat As String: cat = SelectedCategoryName
    Dim cur As Collection: Set cur = LoadAllTasks_File_Cat(cat)         ' ���� ��ü
    Dim add As Collection: Set add = TasksFromListBox(Me.lstTask)       ' ���� ���

    ' 1) �ߺ� ����(�̸�/From/To ���Ͻ� �ߺ�)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long, t As clsTaskItem, k As String
    ' ���� �� dict
    For i = 1 To cur.Count
        Set t = cur(i)
        k = TripleKey(t)
        If Not dict.Exists(k) Then dict.add k, t
    Next
    ' �߰� �� dict (���� ����)
    For i = 1 To add.Count
        Set t = add(i)
        k = TripleKey(t)
        If Not dict.Exists(k) Then dict.add k, t
    Next

    ' 2) dict �� Collection
    Dim merged As New Collection, v
    For Each v In dict.items
        merged.add v
    Next

    ' 3) ����(From �� To �� Name)
    Set merged = SortTasksByFromToName(merged)

    ' 4) ����(ī�װ� ���� + �α�)
    SaveAllTasks_File_Cat cat, merged

    RefreshTaskListAndOverlay
    MsgBox "�߰� ���� �Ϸ�.", vbInformation
    Exit Sub
EH:
    MsgBox "�߰� ���� �� ����: " & Err.Description, vbExclamation
End Sub


' (�̸�|From|To) Ű ? �̸� ��ҹ��� ����, ��¥�� yyyy-mm-dd ����
Private Function TaskKey3(ByVal t As clsTaskItem) As String
    Dim nm As String: nm = LCase$(Trim$(NzCStr(t.TaskName)))
    Dim f As String:  f = Format$(t.FromDate, "yyyy-mm-dd")
    Dim z As String
    If t.HasTo Then
        z = Format$(t.ToDate, "yyyy-mm-dd")
    Else
        z = ""                             ' To ������ �� ���ڿ��� ��
    End If
    TaskKey3 = nm & "|" & f & "|" & z
End Function

' To�� ������ ���� �񱳿� To�� From���� ���
Private Function ToForSort(ByVal t As clsTaskItem) As Date
    If t.HasTo Then ToForSort = t.ToDate Else ToForSort = t.FromDate
End Function

' a<b:-1, a>b:1, ����:0 (From �� To �� Name)
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

' ���� ����: 0/1�� ���� ó�� + ��������
' ����: From �� To �� Name
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

    ' ���� ���� (j ��� ���� Ȯ�� �� �� ���� ��)
    For i = 2 To n
        Dim cur As clsTaskItem
        Set cur = arr(i)
        j = i - 1

        Do While j >= 1
            ' j�� 1 �̸��̸� ������ ����(�ܶ��� ��ü)
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


'=== �߰� ===
Private Sub btnTaskAdd_Click()
    Dim sName As String: sName = Trim$(Me.txtTaskName.text)
    Dim sFrom As String: sFrom = Trim$(Me.txtTaskFrom.text)
    Dim sTo   As String: sTo = Trim$(Me.txtTaskTo.text)

    Dim dF As Date, dT As Date
    Dim okF As Boolean, okT As Boolean

    okF = TryParseDate(sFrom, dF)
    okT = (Len(sTo) > 0 And TryParseDate(sTo, dT))

    If Not okF Then
        MsgBox "������(From)�� �ùٸ��� �ʽ��ϴ�. (yyyy-mm-dd)", vbExclamation
        Exit Sub
    End If
    If okT And dT < dF Then
        MsgBox "������(To)�� �����Ϻ��� ���� �� �����ϴ�.", vbExclamation
        Exit Sub
    End If

    ' �񱳴� ����ȭ�� ���ڿ��� ����
    Dim fromYMD As String, toYMD As String
    fromYMD = Format$(dF, "yyyy-mm-dd")
    toYMD = IIf(okT, Format$(dT, "yyyy-mm-dd"), "")

    ' �ߺ� �˻�
    Dim dupIdx As Long
    dupIdx = FindTaskRow(Me.lstTask, sName, fromYMD, toYMD)

    If dupIdx >= 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "������ �׸��� �̹� �����մϴ�." & vbCrLf & _
            "Task: " & IIf(Len(sName) = 0, "(�̸� ����)", sName) & vbCrLf & _
            "From: " & fromYMD & vbCrLf & _
            "To  : " & IIf(Len(toYMD) = 0, "(����)", toYMD) & vbCrLf & vbCrLf & _
            "�׷��� �߰��Ͻðڽ��ϱ�?", _
            vbExclamation + vbYesNo, "�ߺ� Ȯ��")
        If resp = vbNo Then Exit Sub
    End If

    ' �߰�
    Me.lstTask.AddItem
    Me.lstTask.list(Me.lstTask.ListCount - 1, 0) = sName
    Me.lstTask.list(Me.lstTask.ListCount - 1, 1) = fromYMD
    Me.lstTask.list(Me.lstTask.ListCount - 1, 2) = toYMD

    ' (�ɼ�) �������� ���� ������ ��� ����
    ' RefreshTaskOverlayIfOn
End Sub

'=== ����(������Ʈ) ===
Private Sub btnTaskUpdate_Click()
    Dim r As Long: r = Me.lstTask.ListIndex
    If r < 0 Then
        MsgBox "������ �׸��� �����ϼ���.", vbExclamation
        Exit Sub
    End If

    Dim sName As String: sName = Trim$(Me.txtTaskName.text)
    Dim sFrom As String: sFrom = Trim$(Me.txtTaskFrom.text)
    Dim sTo As String:   sTo = Trim$(Me.txtTaskTo.text)

    Dim dF As Date, dT As Date, okF As Boolean, okT As Boolean
    okF = TryParseDate(sFrom, dF)
    okT = TryParseDate(sTo, dT)

    If Not okF Then
        MsgBox "������(From)�� �ùٸ��� �ʽ��ϴ�. (yyyy-mm-dd)", vbExclamation
        Exit Sub
    End If
    If okT And dT < dF Then
        MsgBox "������(To)�� �����Ϻ��� ���� �� �����ϴ�.", vbExclamation
        Exit Sub
    End If

    Me.lstTask.list(r, 0) = sName
    Me.lstTask.list(r, 1) = Format$(dF, "yyyy-mm-dd")
    Me.lstTask.list(r, 2) = IIf(okT, Format$(dT, "yyyy-mm-dd"), "")
End Sub

'=== ���� ===
Private Sub btnTaskDelete_Click()
    Dim r As Long: r = Me.lstTask.ListIndex
    If r < 0 Then
        MsgBox "������ �׸��� �����ϼ���.", vbExclamation
        Exit Sub
    End If
    Me.lstTask.RemoveItem r
End Sub

'=== ���� (From ����) ===
Private Sub btnTaskSort_Click()
    Dim tasks As Collection
    Set tasks = TasksFromListBox(Me.lstTask)
    Set tasks = SortTasksByFromDate(tasks)
    FillListBoxFromTasksSafe Me.lstTask, tasks
End Sub

'=== ��Ʈ���� �ҷ����� (���ÿ��� 3�� x n��) ===
Private Sub btnTaskGet_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "3�� x n���� ������ ������ �� �����ϼ���.", vbExclamation
        Exit Sub
    End If
    If rng.Columns.Count < 3 Then
        MsgBox "���� ������ �ּ� 3�� ���̾�� �մϴ�. (Name, From, To)", vbExclamation
        Exit Sub
    End If

    Dim r As Range
    For Each r In rng.Rows
        Dim sName As String, sFrom As String, sTo As String
        sName = Trim$(NzCStr(r.Cells(1, 1).Value))
        sFrom = Trim$(NzCStr(r.Cells(1, 2).Value))
        sTo = Trim$(NzCStr(r.Cells(1, 3).Value))

        If Len(sFrom) > 0 Then
            ' ��ȿ���� �߰� ��ư�� ���� ��Ģ ���
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

'=== ��Ʈ�� ��� (New Sheet) ===
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
    MsgBox "�� ��Ʈ�� ����߽��ϴ�: " & ws.name, vbInformation
End Sub

'=== �������� üũ�ڽ� ===
Private Sub chkTaskOverlay_Click()
    mTaskOverlay = (Me.chkTaskOverlay.Value = True)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_OVERLAY, mTaskOverlay
    
    ReapplyOverlayAndSelection (False)

'    ' �ֽ� �׸��� ����
'    BuildDayBoxMapFromGrid
'
'    ' 1) �׻� ���� ���� �ʱ�ȭ
'    ClearAllDayBoxOverlay
'
'    ' ���� ī�װ� ������ Task �ε�
'    Dim tasksAll As Collection
'    Dim s As Date, e As Date
'    Dim cat As String: cat = CStr(Me.cmbTaskCategory.Value)
'    If Me.chkTaskShowAll.Value Then
'        Set tasksAll = LoadAllTasks_File_Cat(cat)                ' ��ü
'    Else
'        GetVisibleDateRange s, e
'        'Set tasksAll = LoadTasksForDateRange_File_Cat(cat, s, e) ' ���̴� ����
'        Set tasksAll = LoadTasksForDateRange_File_Cat(s, e, cat) ' ���̴� ����
'    End If
'
'    ' 2) ON�̸� �� Task�� �ٽ� ����
'    If mTaskOverlay Then
'        ApplyTaskOverlay tasksAll
'    End If
'
'    ' ���� From~To�� �׻� ���� ���ϰ� ������ �ϹǷ� �������� �ٽ� ĥ��
'    'PaintRange
'    ApplySelectedRangeOverlay
    
End Sub
'=== ���ÿ��� üũ�ڽ� ===
Private Sub chkTaskLinkSel_Click()
    mTaskLinkSel = (Me.chkTaskLinkSel.Value = True)
    SaveRegBool REG_SEC_TASKPANEL, REG_KEY_TASK_LINKSEL, mTaskLinkSel
End Sub

'=== ����Ʈ ���� �� From~To �ݿ� ===
' ListBox �̺�Ʈ ����
Private Sub lstTask_Click()
    HandleTaskListSelectionChanged
End Sub

Private Sub lstTask_Change()
    HandleTaskListSelectionChanged
End Sub

' frmYearCalendar ���� �߰�
Private Sub lstTask_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH

    Dim r As Long
    r = Me.lstTask.ListIndex
    If r < 0 Then Exit Sub                      ' ���� ����

    Dim sFrom As String
    sFrom = NzCStr(Me.lstTask.list(r, 1))       ' 0:Name, 1:From, 2:To
    If Len(Trim$(sFrom)) = 0 Then Exit Sub

    Dim dF As Date
    If TryParseYMD(sFrom, dF) Or TryParseDate(sFrom, dF) Then
        ' ���� ���� + ȭ�� �緻��
        SetBaseYear Year(dF)
    Else
        MsgBox "���� �׸��� From ��¥�� �ؼ��� �� �����ϴ�: " & sFrom, vbExclamation
    End If
    Exit Sub
EH:
    ' �ʿ�� �α� ������
End Sub


'=== ��¥�� DayBox ã�� (����) ===
Private Function FindDayBoxByDate(ByVal d As Date) As MSForms.TextBox
    Dim key As String: key = Format$(d, "yyyy-mm-dd")
    If Not DayBoxByDate Is Nothing Then
        If DayBoxByDate.Exists(key) Then
            Set FindDayBoxByDate = DayBoxByDate(key)
        End If
    End If
End Function

'=== �������� ����: ���� �׵θ�/���� ���� (����) ===
Private Sub ClearTaskOverlay()
    ClearAllDayBoxOverlay
    ' ���� ���� ���̶���Ʈ�� ���
    ApplySelectedRangeOverlay
End Sub


'=== �������� ����: �Ķ� �׵θ� + ToolTip ���� (����) ===
'=== �������� ����: �׵θ� + "��� ƾƮ" + ToolTip ���� ===

Private Sub ApplyTaskOverlay(ByVal tasks As Collection)
    ' �⺻/���� ���·� �ʱ�ȭ + �׵θ� ����
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
                ' �׵θ�(�� ����)
                If tb.BorderStyle = 0 Then
                    tb.BorderStyle = fmBorderStyleSingle
                    tb.BorderColor = edge
                End If

                ' ���� ������ ���û��� �ֿ켱 �� ����� �ٲ��� ����
                If Not (haveSel And d >= sSel And d <= eSel And _
                        Not (excl And (IsWeekend(d) Or IsHoliday(d)))) Then
                    ' �������� ���� ��� ����
                    tb.BackColor = fill
                End If

                ' ���� ����
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

' �� ��� �� �ƹ� ��
Private Sub ClearAllDayBoxBorders()
    Dim m As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox
    For m = LBound(tbDay, 1) To UBound(tbDay, 1)
        For r = LBound(tbDay, 2) To UBound(tbDay, 2)
            For c = LBound(tbDay, 3) To UBound(tbDay, 3)
                Set tb = tbDay(m, r, c)
                If Not tb Is Nothing Then
                    On Error Resume Next
                    tb.BorderStyle = 0     ' �� �⺻: ����
                    On Error GoTo 0
                End If
            Next c
        Next r
    Next m
End Sub

'----- Task Color

Private Function TaskPalette() As Variant
    ' ���� �� ��� ���� ���еǴ� 12��(�ʿ�� �� �߰�)

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

' �÷��� ����, �����÷ο�/����ȯ �̽� ���� �ؽ�
' ��ȯ: 0 .. 2,147,483,647  (31-bit ���)
Private Function HashString(ByVal s As String) As Long
    Dim i As Long, ch As Long
    Dim h As Double: h = 5381#                 ' Double ����
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))               ' �����ڵ� �ڵ�����Ʈ(16bit)
        If ch < 0 Then ch = ch + 65536         ' ����ȣ 16bit�� ����ȭ
        h = h * 33# + ch                       ' DJB2 ������ (��Ʈ���� ����)
        ' 2^31 �� ��ⷯ: 0 <= h < 2^31 �� �ǵ��� ����
        h = h - 2147483648# * Fix(h / 2147483648#)
    Next
    HashString = CLng(h)                       ' ����: 0..2147483647 ����
End Function

Private Function MakeTaskKey(ByVal t As clsTaskItem) As String
    ' �̸��� ������ �̸�����, ������ �Ⱓ ���ڿ��� Ű ����(�� �ϰ��� ���� ����)
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
    clr = CLng(s)           ' ������ �׳� Long ���ڿ��� ������ �״�� ����
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
    ' ĳ�á淹����Ʈ�����ȷ�Ʈ �ؽ� ������ ���� ����
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

    ' ������Ʈ���� ������ ���� �ȷ�Ʈ���� �ؽ÷� ����
    Dim p As Variant: p = TaskPalette()
    Dim idx As Long: idx = HashString(UCase$(key)) Mod (UBound(p) - LBound(p) + 1)
    c = p(LBound(p) + idx)

    TaskColorByName.add key, c
    SaveTaskColorToReg key, c
    ColorForTaskKey = c
End Function

Private Sub EnsureColorsForTasks(ByVal tasks As Collection)
    ' �̸� �� �� ������ ���� �� �غ�(������������ ��������)
    If tasks Is Nothing Then Exit Sub
    Dim i As Long, k As String
    For i = 1 To tasks.Count
        k = MakeTaskKey(tasks(i))
        Call ColorForTaskKey(k)
    Next
End Sub

' lstTask�� ���� �׸��� ������ �� �ε���(0-base), ������ -1
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
        MsgBox "��ȿ�� ������ �ƴմϴ�: " & CStr(y), vbExclamation
        Exit Sub
    End If
    Dim cat As String: cat = SelectedCategoryName

    Dim resp As VbMsgBoxResult
    resp = MsgBox("ī�װ� '" & cat & "'���� ���� " & y & "�� Task�� ��� �����մϴ�." & vbCrLf & _
                  "�ǵ��� �� �����ϴ�. �����Ͻðڽ��ϱ�?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "������ Task ���� Ȯ��")
    If resp <> vbYes Then Exit Sub

    On Error GoTo EH
    RemoveTasksForYear_FromAll_Cat cat, y
    RefreshTaskListAndOverlay
    MsgBox "���� " & y & "�� Task�� �����Ǿ����ϴ�.", vbInformation
    Exit Sub
EH:
    MsgBox "���� �� ����: " & Err.Description, vbExclamation
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
' ���̴� 12�� ���� ù°��~��������
Private Sub GetVisibleDateRange(ByRef dStart As Date, ByRef dEnd As Date)
    Dim baseYear As Long: baseYear = mYear
    Dim curM As Long: curM = Month(Date)
    Dim mapDate() As Date
    BuildMonthMap baseYear, curM, mLayoutMode, mapDate, mAnchorSlot
    dStart = DateSerial(Year(mapDate(1)), Month(mapDate(1)), 1)
    dEnd = DateSerial(Year(mapDate(12)), Month(mapDate(12)) + 1, 0)
End Sub

' ���̴� �޷� ������ ��� �������� Task �ε� �� �̸�/�Ⱓ ���� �� ����Ʈ/�������� ����
Private Sub LoadTasksForVisibleRangeAndOverlay()
    Dim s As Date, e As Date
    GetVisibleDateRange s, e

    Dim col As Collection
    Set col = LoadTasksForDateRange_File(s, e)   ' �� JSON ���Ͽ��� ���� �ε� + ����

    FillListBoxFromTasksSafe Me.lstTask, col

    BuildDayBoxMapFromGrid
    If mTaskOverlay Then
        ApplyTaskOverlay col
    Else
        ClearTaskOverlay
    End If
End Sub

' �뷮 �����͸� ������ ä��(.AddItem ���� ��� 2D �迭�� .List ����)
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

    ' �� �ε� (ǥ�ÿ�)
    If Me.chkTaskShowAll.Value Then
        Set tasks = LoadAllTasks_File_Cat(cat)                  ' ǥ�ÿ�(��ü)
    Else
        GetVisibleDateRange s, e
        Set tasks = LoadTasksForDateRange_File_Cat(s, e, cat)   ' ǥ�ÿ�(����)
    End If

    ' ����Ʈ���� �̸� ����(txtTaskFilter)�� ����
    Set displayTasks = FilterTasksByNameLike(tasks, NzCStr(Me.txtTaskFilter.text))

    FillListBoxFromTasksSafe Me.lstTask, displayTasks  ' (���� ���� ��� ������� ����)

    ' �� �������̴� "��� Task" �������� (��� ����)
    '     - ShowAll�̸� tasks �״��, �ƴϸ� ���� ����(�ʿ� �� ��� ���� �ε�� �ٲ㵵 ����)
    Set tasksForOverlay = displayTasks

    ' �� DayBox ���� �ֽ�ȭ
    BuildDayBoxMapFromGrid
    
    ' �ٽ�: ���� ���� ���̽��� �ʱ�ȭ(������ Task �̸� ���� ����)
    ClearAllDayBoxOverlay

    ' �� �������̰� ���� ������ �̹� ī�װ� Task�� �ٽ� ����
    If mTaskOverlay Then
        ApplyTaskOverlay tasksForOverlay
    End If

    ' �� ���� From~To�� �׻� �ֻ����� ���̵��� �������� �ٽ� ĥ��
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

' ��� ���� �ະ�� ��� ���� (IIf ��� ����)
Private Sub FillListBoxFromTasksSafe(lst As MSForms.ListBox, ByVal tasks As Collection)
    Dim n As Long: n = IIf(tasks Is Nothing, 0, tasks.Count)
    Dim data As Variant
    Dim i As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnHeads = False
    ' �ʿ� �� ������: lst.ColumnWidths = "140 pt;75 pt;75 pt"
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

' ���� �׸�� ������ (Name, From, To) Task�� JSON ����ҿ��� ����
Private Sub btnTaskDeleteSelected_Click()
    Dim sel As Collection: Set sel = SelectedTasksFromList(Me.lstTask)
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "������ �׸��� �����ϼ���.", vbExclamation
        Exit Sub
    End If

    Dim cat As String: cat = SelectedCategoryName
    Dim all As Collection: Set all = LoadAllTasks_File_Cat(cat)

    ' ���� Ű ����
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    Dim i As Long, t As clsTaskItem
    For i = 1 To sel.Count
        keys(TripleKey(sel(i))) = True
    Next

    ' ���͸�
    Dim remain As New Collection
    For i = 1 To all.Count
        Set t = all(i)
        If Not keys.Exists(TripleKey(t)) Then remain.add t
    Next

    On Error GoTo EH
    SaveAllTasks_File_Cat cat, remain
    RefreshTaskListAndOverlay
    MsgBox "���� �׸��� �����߽��ϴ�.", vbInformation
    Exit Sub
EH:
    MsgBox "���� �� ����: " & Err.Description, vbExclamation
End Sub

' name|yyyy-mm-dd|yyyy-mm-dd(�Ǵ� ���ڿ�) �� ��Ű ����
Private Function TripleKey(ByVal t As clsTaskItem) As String
    Dim f As String, z As String
    f = Format$(t.FromDate, "yyyy-mm-dd")
    z = IIf(t.HasTo, Format$(t.ToDate, "yyyy-mm-dd"), "")
    TripleKey = NzCStr(t.TaskName) & "|" & f & "|" & z
End Function

' (name, from, to) �� �񱳿� Ű (�̸��� ��ҹ��� ����)
Private Function MakeTripleKey(ByVal nm As String, ByVal fromYMD As String, ByVal toYMD As String) As String
    MakeTripleKey = LCase$(Trim$(NzCStr(nm))) & "|" & Trim$(NzCStr(fromYMD)) & "|" & Trim$(NzCStr(toYMD))
End Function

' clsTaskItem �� �񱳿� Ű
Private Function TaskTripleKey(ByVal t As clsTaskItem) As String
    TaskTripleKey = MakeTripleKey( _
        t.TaskName, _
        Format$(t.FromDate, "yyyy-mm-dd"), _
        IIf(t.HasTo, Format$(t.ToDate, "yyyy-mm-dd"), "") _
    )
End Function

' lstTask���� ���õ� ��� �׸��� Ű ����(Dictionary) ��ȯ
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


' ListBox�� ���� ���� �� ȣ��: ���õ� Task�� �������� + From~To = Min~Max
Private Sub HandleTaskListSelectionChanged()

    ' === ��ũ ������ ���� ������ �ƹ� �͵� �� �� ===
    If Not (Me.chkTaskLinkSel.Value = True) Then Exit Sub
    
    Dim sel As Collection: Set sel = SelectedTasksFromList(Me.lstTask)

    If sel Is Nothing Or sel.Count = 0 Then
        ' ���� ���� ���� ���� ���� From~To �������� ����
        Set mSelRanges = Nothing
        PaintRange                      ' ���� ���� ���� ĥ�ϱ�
        ' �������̴� ������
        Exit Sub
    End If

    ' 1) ���� ���� ���
    Set mSelRanges = BuildRangesFromTasks(sel)

    ' 2) �ּڰ�/�ִ����� From~To �ؽ�Ʈ �� ���� ���� ����
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

    ' 3) ���õ� �����鸸 ��� ĥ�ϱ� (�������� �׵θ��� �ǵ帮�� ����)
    PaintSelectionRangesMulti mSelRanges

    ' (����) �������̰� ���� �ִٸ� �׵θ��� �ֽ� ȭ�鿡 �������ϰ� ���� ��:
    If mTaskOverlay Then RefreshTaskOverlayIfOn
    
    Dim selCount As Long, firstIdx As Long
    firstIdx = -1

    For i = 0 To Me.lstTask.ListCount - 1
        If Me.lstTask.Selected(i) Then
            selCount = selCount + 1
            If firstIdx < 0 Then firstIdx = i
            If selCount > 1 Then Exit For  ' ���� �����̸� �ؽ�Ʈ�ڽ� ���� �� ��
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

' ���õ� ��鸸 clsTaskItem���� ��ȯ
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

' Ư�� ���� clsTaskItem���� ��ȯ
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

' ���õ� Task��κ��� [����~��] ���� �迭 ����
Private Function BuildRangesFromTasks(ByVal tasks As Collection) As Collection
    Dim rngs As New Collection, i As Long, s As Date, e As Date, tmp As Date
    If Not tasks Is Nothing Then
        For i = 1 To tasks.Count
            s = tasks(i).FromDate
            e = IIf(tasks(i).HasTo, tasks(i).ToDate, s)
            If e < s Then tmp = s: s = e: e = tmp
            rngs.add Array(s, e)   ' Variant(2) ���
        Next
    End If
    Set BuildRangesFromTasks = rngs
End Function

' ��¥�� ������ ������ �ȿ� ���ԵǴ°�?
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

' ���� ������ �����ÿ����� ������ ĥ�ϱ� (Border�� �ǵ帮�� ���� �� �������� ����)
Private Sub PaintSelectionRangesMulti(ByVal ranges As Collection)
    Dim mBlock As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox, sTag As String, dT As Date
    Dim excludeNonBiz As Boolean: excludeNonBiz = chkExcludeNonBiz.Value

    For mBlock = 1 To 12
        For r = 1 To 6
            For c = 1 To 7
                Set tb = tbDay(mBlock, r, c)
                sTag = Trim$(tb.Tag)

                ' �⺻ ��Ÿ�� ���� (���/����/����/���� ��) - Border�� ������ ����
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

    ' �⺻ ī�װ� "tasks"�� �׻� ����
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

    ' ������ ���� ����(������ tasks)
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
    RefreshTaskListAndOverlay    ' ī�װ� �ٲ�� ��� �ٽ� �ε�/��������
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
    nm = InputBox("�߰��� �������� �Է��ϼ���.", "���� �߰�")
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Sub

    nm = SafeFileBaseName(nm)

    ' �̹� ��Ͽ� ������ ���ø�
    If SelectComboValue(Me.cmbTaskCategory, nm) Then Exit Sub

    ' �� ���� ����([]) + �α� ������
    Dim emptyCol As New Collection
    SaveAllTasks_File_Cat nm, emptyCol

    ' �޺� ���� �� ����
    LoadCategoryList
    SelectComboValue Me.cmbTaskCategory, nm
End Sub

Private Sub btnTaskCategoryDelete_Click()
    Dim cat As String: cat = SelectedCategoryName
    If LCase$(cat) = "tasks" Then
        MsgBox "�⺻ ī�װ� 'tasks'�� ������ �� �����ϴ�.", vbExclamation
        Exit Sub
    End If

    Dim r As VbMsgBoxResult
    r = MsgBox("���� '" & cat & "'�� JSON ������ �����մϴ�." & vbCrLf & _
               "�α� ������ �ش� �������� �Բ� �����ұ��?", _
               vbQuestion + vbYesNoCancel, "���� ����")
    If r = vbCancel Then Exit Sub

    On Error Resume Next
    ' �� ���� ����
    RemoveAllTasks_File_Cat cat

    ' �α� ���� ����(vbYes�� �α׵� ����)
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


' name�� filterRaw(��: "dev,release*") �� �ϳ��� ��Ī�Ǹ� True
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
            ' ���ϵ�ī�尡 ������ �κ���ġ�� ��ȯ
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

' �÷��ǿ��� TaskName �������� ����
Private Function FilterTasksByNameLike(ByVal tasks As Collection, ByVal filterRaw As String) As Collection
    Dim out As New Collection, i As Long, t As clsTaskItem, nm As String
    Dim raw As String: raw = Trim$(filterRaw)
    If raw = "" Or raw = "*" Then
        ' ���� ��Ȱ��: ���� �״�� ��ȯ(���� ��ȯ�̶� ���ɡ�)
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

' �� ���(frmYearCalendar)
Private Function CurrentCategoryName() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(Me.cmbTaskCategory.Value)
    On Error GoTo 0
    If Len(s) = 0 Then s = "tasks"   ' �⺻ ī�װ�
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

' === �� ����� ===
Private Function MixColors(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    ' ��� = c1*(1-t) + c2*t  (t: 0~1)
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
    ' �׵θ����� ��� ������ 70% ��� �� ������ ĥ(�������� ����)
    OverlayFillColorFor = MixColors(edgeColor, RGB(255, 255, 255), 0.7)
End Function


' �� ���(frmYearCalendar) ��򰡿� �߰�
Private Sub PaintSelectedCell(ByVal tb As MSForms.TextBox)
    With tb
        .BackColor = RGB(33, 92, 152)   ' ���� �Ķ�
        .ForeColor = RGB(255, 255, 255) ' �� ����
        .Font.Bold = True
        '���ϸ� �׵θ��� ����
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(33, 92, 152)
    End With
End Sub

'=== REPLACE: ��������/���� ���� �ʱ�ȭ(�ָ�/������/���� �� ���� + �׵θ� ���� + ����=��¥/�����ϸ�) ===
Private Sub ClearAllDayBoxOverlay()
    Dim m As Long, r As Long, c As Long
    Dim tb As MSForms.TextBox
    Dim sTag As String, d As Date

    For m = LBound(tbDay, 1) To UBound(tbDay, 1)
        For r = LBound(tbDay, 2) To UBound(tbDay, 2)
            For c = LBound(tbDay, 3) To UBound(tbDay, 3)
                Set tb = tbDay(m, r, c)
                If Not tb Is Nothing Then
                    ' 1) �⺻ �����۲� ����
                    ApplyBaseStyle tb, c
                    ' 2) �׵θ� ����
                    On Error Resume Next
                    tb.BorderStyle = 0
                    On Error GoTo 0
                    ' 3) ������ '��¥(+������)'������ �缳��
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

'=== NEW: ��¥ + (������) �����ϸ����� ���̽� ���� ���� ===
Private Function BaseTipForDate(ByVal d As Date) As String
    Dim tip As String, holName As String, holPrefix As String
    tip = FmtDateOut(d)                                ' ��¥ 1��
    holName = GetHolidayNameIfAny(d)                   ' ������ �̸�
    If Len(holName) > 0 Then
        holPrefix = IIf(mLang = LangE, "[Holiday] ", "[������] ")
        tip = tip & vbCrLf & holPrefix & holName
    End If
    BaseTipForDate = tip
End Function

' ���� ������ �׻� ���� �������� ������ ĥ�� ���̶���Ʈ
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
        ' �� ���� �ɼ��� ���� ������ ��/��/�������� �ǳʶ�
        If excludeNonBiz Then
            If IsWeekend(d) Or IsHoliday(d) Then GoTo ContinueNextDate
        End If

        Dim tb As MSForms.TextBox
        Set tb = FindDayBoxByDate(d)
        If Not tb Is Nothing Then
            ' ���� ������ �ֿ켱(��������/�ָ��� �� ���� ���)
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
        RenderAllMonths              ' �޷� �� ����/�⺻ ��Ÿ��
        BuildDayBoxMapFromGrid
    End If

    ClearTaskOverlay                ' �׵θ�/����/��� ����(���� ǥ�� X)
    If mTaskOverlay Then
        ApplyTaskOverlay TasksFromListBox(Me.lstTask)  ' ���/�׵θ�/����
    End If

    PaintRange                      ' �� �������� ���ù��� ĥ�ϱ�(���� ���ϰ�)

LFinally:
    mRefreshing = False
End Sub

Private Sub ReapplyOverlayAndSelection(Optional ByVal rerender As Boolean = False)
    RefreshVisuals rerender                               ' ��3 ���ɽ�Ʈ������
End Sub

Private Function SortTasksByFromDate(ByVal tasks As Collection) As Collection
    ' �ܼ� ��������: �迭�� �Ű� ���� �� �����
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
        Set t = New clsTaskItem                 ' �ڡ� �� �ݺ����� New �ʼ� �ڡ�
        
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
            out.add t                              ' �� �� �ν��Ͻ��� �߰�
        End If
    Next

    Set TasksFromListBox = out
End Function

