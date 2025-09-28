Attribute VB_Name = "modYearCalHelp"
' �� 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.

Option Explicit

' ���� ���: ���� ���(K/E) ��ȸ
Private Function GetCurrentLangFromRegistry() As String
    ' frmYearCalendar�� ���� Ű�� �����ϰ� ����
    Dim appName As String, secI18N As String
    appName = "PeriodPicker"
    secI18N = "I18N"
    GetCurrentLangFromRegistry = UCase$(GetSetting(appName, secI18N, "Lang", "K"))
End Function

' (ȣȯ��) ���� ���� ȣ���ϴ� �̸� ����
' �� ���������� ���� �� �����ؼ� ��/�� �Ŵ����� ��ȯ
Public Function BuildYearCalendarManual() As String
    BuildYearCalendarManual = BuildYearCalendarManual_Current()
End Function

' ���� ��� �������� ��ȯ(K/E)
Public Function BuildYearCalendarManual_Current() As String
    If GetCurrentLangFromRegistry() = "E" Then
        BuildYearCalendarManual_Current = BuildYearCalendarManual_EN()
    Else
        BuildYearCalendarManual_Current = BuildYearCalendarManual_KO()
    End If
End Function

' ��������������������������������������������������������������������������������������������������������������������������
' �ѱ��� �Ŵ���
' ��������������������������������������������������������������������������������������������������������������������������
Public Function BuildYearCalendarManual_KO() As String
    Dim t As String, NL As String: NL = vbCrLf

    t = t & "�� Year Calendar - ����" & NL
    t = t & "����������������������������������������������������������������" & NL & NL

    t = t & "�� ����" & NL
    t = t & "- 1�� �޷��� 4��3 �׸���� ǥ���ϰ�, �Ⱓ(From~To)�� ������ ����/��ȯ�մϴ�." & NL
    t = t & "- ���/��¥����/���Ͻ���/���ϸ� ��Ÿ���� �������� ���� �����ϸ� ������Ʈ���� ����˴ϴ�." & NL & NL

    t = t & "�� �⺻ ����" & NL
    t = t & "- ���� �� ��Ŭ��=From, ��Ŭ��=To, ����Ŭ��=��ù�ȯ(���� ����)" & NL
    t = t & "- �� Ÿ��Ʋ ��Ŭ��=�ش�� 1��(From), ��Ŭ��=����(To), ����Ŭ��=�� �� ��ü" & NL
    t = t & "- [����]���� ��ȯ, [Clear]�� �ʱ�ȭ, [�ݱ�]�� ����" & NL & NL

    t = t & "�� Ű����/��" & NL
    t = t & "- PageUp/Down: ��1�� | Shift: ��10�� | Ctrl: ��5�� | Shift+Ctrl: ��15��" & NL
    t = t & "- End: ������ ������ �̵�, Enter: ����, Esc: �ݱ�, Del: �ʱ�ȭ" & NL
    t = t & "- ���콺��: ���� ��1 (Shift/Ctrl ������ ���� ���� ���)" & NL
    t = t & "- ��ǥ(,): From=����  |  ��ħǥ(.): To=����" & NL & NL

    t = t & "�� ���̾ƿ�" & NL
    t = t & "- 1��12, ������� ����/����, ����� @Slot(1~12) ��� ����" & NL
    t = t & "- ����� ����� ���� �Ķ� ������� ���̶���Ʈ" & NL & NL

    t = t & "�� ����" & NL
    t = t & "- ���(K/E), ��¥����, �� Ÿ��Ʋ ����(��), ���� ����(��/��), ���ϸ� ��Ÿ��(��Ī/Ǯ����)" & NL
    t = t & "- ������ �� Ÿ��Ʋ ����(���� ��ū ����)�� �⺻������ �ڵ� ����" & NL & NL

    t = t & "�� ���� ���̶���Ʈ/������" & NL
    t = t & "- ���� ĥ�ϱ�� �����, [����ǥ�ÿ� ���� ����] ���� �� ��/��/������ ����" & NL
    t = t & "- �ϴ� ���: (������ / ���ϼ�), �������� gHolidaySet ���" & NL & NL

    t = t & "�� Export" & NL
    t = t & "- �� ��Ʈ�� ���̾ƿ�/����/���ù��� ���� �Բ� �޷��� ���" & NL
    t = t & "- ������ �� �ڸ�Ʈ�� �̸� ǥ��, ����/�ָ�/���� �� ����" & NL & NL

'    t = t & "�� ������Ʈ��(���� ��ġ)" & NL
'    t = t & "- App: MonthlyCalendar" & NL
'    t = t & "  �� I18N\\Lang, I18N\\WeekStart, I18N\\WeekNameStyle" & NL
'    t = t & "  �� Format\\Date, Format\\MonthTitle_K/E" & NL
'    t = t & "  �� Layout\\LayoutMode/AnchorSlot/Apply_From_Only/Exclude_Non_Biz/Show_Keep" & NL & NL

    t = t & "�� ��/�˸�" & NL
    t = t & "- ��ǥ/��ħǥ Ű�� VK �ڵ�(188/190) ��� ����(ȯ�溰 ���� ����)" & NL
    t = t & "- ������ ���ۿ����� �Ͽ��ϸ� ����, ����ϸ� �Ķ����� ǥ�õ˴ϴ�." & NL

    BuildYearCalendarManual_KO = t
End Function

' ��������������������������������������������������������������������������������������������������������������������������
' English manual
' ��������������������������������������������������������������������������������������������������������������������������
Public Function BuildYearCalendarManual_EN() As String
    Dim t As String, NL As String: NL = vbCrLf

    t = t & "�� Year Calendar - User Guide" & NL
    t = t & "����������������������������������������������������������������" & NL & NL

    t = t & "�� Purpose" & NL
    t = t & "- Display a full-year calendar in a 4��3 grid and quickly select a date range (From?To)." & NL
    t = t & "- Language/date format/week start/day name style are configurable and persisted in the registry." & NL & NL

    t = t & "�� Basic Operations" & NL
    t = t & "- Day cell Left-Click = From, Right-Click = To, Double-Click = return immediately (auto-order)." & NL
    t = t & "- Month title Left-Click = 1st day (From), Right-Click = last day (To), Double-Click = whole month." & NL
    t = t & "- Use [Apply] to return the range, [Clear] to reset, [Close] to exit." & NL & NL

    t = t & "�� Keyboard & Wheel" & NL
    t = t & "- PageUp/Down: ��1 year | Shift: ��10 | Ctrl: ��5 | Shift+Ctrl: ��15" & NL
    t = t & "- End: jump to current year, Enter: Apply, Esc: Close, Del: Reset" & NL
    t = t & "- Mouse wheel: ��1 year (same multipliers with Shift/Ctrl)." & NL
    t = t & "- Comma ( , ): From = Today  |  Period ( . ): To = Today" & NL & NL

    t = t & "�� Layout" & NL
    t = t & "- Modes: 1��12, Current First, Current Last, Current @Slot (1?12)." & NL
    t = t & "- The current month block is gently highlighted in the background." & NL & NL

    t = t & "�� Settings" & NL
    t = t & "- Language (K/E), Date format, Month title format (per language)," & NL
    t = t & "  Week start (Sun/Mon), Day name style (Short/Full)." & NL
    t = t & "- Invalid month-title formats (containing day tokens) are auto-sanitized to defaults." & NL & NL

    t = t & "�� Range Highlight & Business Days" & NL
    t = t & "- Selected range is painted in light yellow; option to exclude Sat/Sun/Holidays from painting." & NL
    t = t & "- Summary shows (Business days / Total days). Holidays rely on gHolidaySet." & NL & NL

    t = t & "�� Export" & NL
    t = t & "- Prints the current layout/year/range summary to a new worksheet." & NL
    t = t & "- Holiday names are inserted as cell comments; weekend/today/range coloring retained." & NL & NL

'    t = t & "�� Registry (Persistence)" & NL
'    t = t & "- App: MonthlyCalendar" & NL
'    t = t & "  �� I18N\\Lang, I18N\\WeekStart, I18N\\WeekNameStyle" & NL
'    t = t & "  �� Format\\Date, Format\\MonthTitle_K/E" & NL
'    t = t & "  �� Layout\\LayoutMode, AnchorSlot, Apply_From_Only, Exclude_Non_Biz, Show_Keep" & NL & NL

    t = t & "�� Tips / Notes" & NL
    t = t & "- For keyboard handling, prefer VK codes for comma/period (188/190) to avoid locale issues." & NL
    t = t & "- With Monday as week start, only Sunday is red and Saturday is blue." & NL

    BuildYearCalendarManual_EN = t
End Function

' ��������������������������������������������������������������������������������������������������������������������������
' (����) ������ ����/���� ������ ���� ���� �� ȣ���� ����
' ��������������������������������������������������������������������������������������������������������������������������
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

