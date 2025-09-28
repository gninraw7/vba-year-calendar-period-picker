Attribute VB_Name = "modCallBackRibbon"
' �� 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
'
'' ���� : �ӱ��� (M:O1O-2716-0735, E-Mail: gninraw7@naver.com, kiljae.lim@gmail.com)
'' 2025-09-08 ���� �ۼ�
'' 2025-09-15 Addin �ۼ�

Option Explicit

' Ribbon �ڵ� & ����
Public gRibbon As IRibbonUI

Public Const Const_PeriodPicker_Menu As String = "PeriodPicker_Menu"

' �̹� ����Ǿ� ���� �ʴٸ� ��� ��� �� ���� ������ ����
Public gHolidaySet As Object   ' 'Scripting.Dictionary'

Private Const REG_APP As String = "PeriodPicker"
Private Const REG_SEC As String = "Holidays"

'----------------------
' �ʱ�ȭ / Ribbon �⺻
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
    f.SetTargetRange Selection         ' ���ÿ����� ������ ���ο��� Selection ����
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
        .Caption = "����"
        .Tag = "PeriodPicker_Cell_Control_Tag"
    End With
    
    With L_CommandBar.Controls.add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "InsertToday"
        .Caption = "����"
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
    ' ���õ� ���� ���� ��¥ �Է�
    If Not Selection Is Nothing Then
        Selection.Value = Date
    End If
End Sub

Sub InsertYesterday()
   ' ���õ� ���� ���� ��¥ �Է�
    If Not Selection Is Nothing Then
        Selection.Value = Date - 1
    End If
End Sub


'���� ������ 1���̸� "YYYY-MM-DD ~ YYYY-MM-DD" ���ڿ��� ä��
'2�� �̻��̸� ù ��=From, �� ������ ��=To(��¥ ����)
'
'(2) �ٸ� ���� TextBox �� ���� ��ȯ
' ��: frmOrder �� TextBox txtFrom, txtTo �� �ִٰ� ����
'Sub ShowYearCalendar_ToOtherFormTextBoxes()
'    Dim f As New frmYearCalendar
'    f.SetTargetTextBoxes frmOrder.txtFrom, frmOrder.txtTo
'    f.Show
'End Sub


' ================== ����: ��� ���� �� gHolidaySet �ε� ==================
Public Sub LoadAllHolidaysIntoGlobalSet()
    ' 1) dict �غ�
    If gHolidaySet Is Nothing Then
        Set gHolidaySet = CreateObject("Scripting.Dictionary")
    Else
        gHolidaySet.RemoveAll
    End If
    ' ��¥ Ű�� ���� ���̹Ƿ� CompareMode ���� ����

    Dim kv As Variant, i As Long, found As Boolean
    On Error Resume Next
    kv = GetAllSettings(REG_APP, REG_SEC)   ' 2���� �迭([row, 0]=key, [row, 1]=value)
    On Error GoTo 0

    If IsArray(kv) Then
        ' 2) ���ǿ� �����ϴ� ��� (����Ű, ����) ó��
        For i = LBound(kv, 1) To UBound(kv, 1)
            AddHolidayRawToDict CStr(kv(i, 1)), gHolidaySet  ' value=����
        Next
        found = True
    End If

    ' 3) ����: ������ ���ų� �б� ���� �� ���� ���� ��ĵ
    If Not found Then
        Dim y As Long, raw As String
        For y = 1900 To 2099
            raw = GetSetting(REG_APP, REG_SEC, CStr(y), "")
            If Len(raw) > 0 Then AddHolidayRawToDict raw, gHolidaySet
        Next
    End If

    ' ���ϸ� ����� ī��Ʈ ���
    'Debug.Print "gHolidaySet.Count=" & gHolidaySet.Count
End Sub

' ================== ���� �Ľ�: "yyyy-mm-dd|���ϸ�" �ٵ� �� dict(Date �� �̸�) ==================
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
            ' ���� ��¥�� �̹� ������ ������ ���� �켱���� ���
            dict(CDate(Int(CDbl(dT)))) = nm   ' �������� ����ȭ(��¥ Ű ����ȭ)
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
            def = DefaultHolidayLines(y)   ' �� �Ʒ� �Լ��� ����� �⺻�� ���
            If Len(def) > 0 Then
                SaveSetting REG_APP, REG_SEC, CStr(y), def
                wrote = wrote + 1
            End If
        End If
    Next

'    MsgBox IIf(forceOverwrite, "[����]", "[����ִ� ������]") & _
'           " ������ �õ� �Ϸ�: " & wrote & "�� ����", vbInformation
End Sub

'=========================
' �� ���� �⺻��(���ȭ�� ����)  �� ���� �ٵ� �������� ��¹����� ��ü�ϼ���
'     ����: ������Ʈ���� ������ �δ� raw ���ڿ� �״��
'     (�ٴ� "YYYY-MM-DD<������>���ϸ�" ����, �����ڴ� ���� ���� ���� ����)
'=========================
Public Function DefaultHolidayLines(ByVal y As Long) As String
    Select Case y
        Case 2004: DefaultHolidayLines = "2004-01-01|����" & vbCrLf & "2004-01-21|����" & vbCrLf & "2004-01-22|����" & vbCrLf & "2004-01-23|����" & vbCrLf & "2004-03-01|������" & vbCrLf & "2004-04-05|�ĸ���" & vbCrLf & "2004-05-05|��̳�" & vbCrLf & "2004-05-26|����ź����" & vbCrLf & "2004-06-06|������" & vbCrLf & "2004-07-17|������" & vbCrLf & "2004-08-15|������" & vbCrLf & "2004-09-27|�߼�" & vbCrLf & "2004-09-28|�߼�" & vbCrLf & "2004-09-29|�߼�" & vbCrLf & "2004-10-03|��õ��" & vbCrLf & "2004-12-25|�⵶ź����" & vbCrLf & ""
        Case 2005: DefaultHolidayLines = "2005-01-01|����" & vbCrLf & "2005-02-08|����" & vbCrLf & "2005-02-09|����" & vbCrLf & "2005-02-10|����" & vbCrLf & "2005-03-01|������" & vbCrLf & "2005-04-05|�ĸ���" & vbCrLf & "2005-05-05|��̳�" & vbCrLf & "2005-05-15|����ź����" & vbCrLf & "2005-06-06|������" & vbCrLf & "2005-07-17|������" & vbCrLf & "2005-08-15|������" & vbCrLf & "2005-09-17|�߼�" & vbCrLf & "2005-09-18|�߼�" & vbCrLf & "2005-09-19|�߼�" & vbCrLf & "2005-10-03|��õ��" & vbCrLf & "2005-12-25|�⵶ź����" & vbCrLf & ""
        Case 2006: DefaultHolidayLines = "2006-01-01|����" & vbCrLf & "2006-01-28|����" & vbCrLf & "2006-01-29|����" & vbCrLf & "2006-01-30|����" & vbCrLf & "2006-03-01|������" & vbCrLf & "2006-05-05|����ź����" & vbCrLf & "2006-06-06|������" & vbCrLf & "2006-07-17|������" & vbCrLf & "2006-08-15|������" & vbCrLf & "2006-10-03|��õ��" & vbCrLf & "2006-10-05|�߼�" & vbCrLf & "2006-10-06|�߼�" & vbCrLf & "2006-10-07|�߼�" & vbCrLf & "2006-12-25|�⵶ź����" & vbCrLf & ""
        Case 2007: DefaultHolidayLines = "2007-01-01|����" & vbCrLf & "2007-02-17|����" & vbCrLf & "2007-02-18|����" & vbCrLf & "2007-02-19|����" & vbCrLf & "2007-03-01|������" & vbCrLf & "2007-05-05|��̳�" & vbCrLf & "2007-05-24|����ź����" & vbCrLf & "2007-06-06|������" & vbCrLf & "2007-07-17|������" & vbCrLf & "2007-08-15|������" & vbCrLf & "2007-09-24|�߼�" & vbCrLf & "2007-09-25|�߼�" & vbCrLf & "2007-09-26|�߼�" & vbCrLf & "2007-10-03|��õ��" & vbCrLf & "2007-12-25|�⵶ź����" & vbCrLf & ""
        Case 2008: DefaultHolidayLines = "2008-01-01|����" & vbCrLf & "2008-02-06|����" & vbCrLf & "2008-02-07|����" & vbCrLf & "2008-02-08|����" & vbCrLf & "2008-03-01|������" & vbCrLf & "2008-05-05|��̳�" & vbCrLf & "2008-05-12|����ź����" & vbCrLf & "2008-06-06|������" & vbCrLf & "2008-08-15|������" & vbCrLf & "2008-09-13|�߼�" & vbCrLf & "2008-09-14|�߼�" & vbCrLf & "2008-09-15|�߼�" & vbCrLf & "2008-10-03|��õ��" & vbCrLf & "2008-12-25|�⵶ź����" & vbCrLf & ""
        Case 2009: DefaultHolidayLines = "2009-01-01|����" & vbCrLf & "2009-01-25|����" & vbCrLf & "2009-01-26|����" & vbCrLf & "2009-01-27|����" & vbCrLf & "2009-03-01|������" & vbCrLf & "2009-05-02|����ź����" & vbCrLf & "2009-05-05|��̳�" & vbCrLf & "2009-06-06|������" & vbCrLf & "2009-08-15|������" & vbCrLf & "2009-10-02|�߼�" & vbCrLf & "2009-10-03|��õ��" & vbCrLf & "2009-10-04|�߼�" & vbCrLf & "2009-12-25|�⵶ź����" & vbCrLf & ""
        Case 2010: DefaultHolidayLines = "2010-01-01|����" & vbCrLf & "2010-02-13|����" & vbCrLf & "2010-02-14|����" & vbCrLf & "2010-02-15|����" & vbCrLf & "2010-03-01|������" & vbCrLf & "2010-05-05|��̳�" & vbCrLf & "2010-05-21|����ź����" & vbCrLf & "2010-06-06|������" & vbCrLf & "2010-08-15|������" & vbCrLf & "2010-09-21|�߼�" & vbCrLf & "2010-09-22|�߼�" & vbCrLf & "2010-09-23|�߼�" & vbCrLf & "2010-10-03|��õ��" & vbCrLf & "2010-12-25|�⵶ź����" & vbCrLf & ""
        Case 2011: DefaultHolidayLines = "2011-01-01|����" & vbCrLf & "2011-02-02|����" & vbCrLf & "2011-02-03|����" & vbCrLf & "2011-02-04|����" & vbCrLf & "2011-03-01|������" & vbCrLf & "2011-05-05|��̳�" & vbCrLf & "2011-05-10|����ź����" & vbCrLf & "2011-06-06|������" & vbCrLf & "2011-08-15|������" & vbCrLf & "2011-09-11|�߼�" & vbCrLf & "2011-09-12|�߼�" & vbCrLf & "2011-09-13|�߼�" & vbCrLf & "2011-10-03|��õ��" & vbCrLf & "2011-12-25|�⵶ź����" & vbCrLf & ""
        Case 2012: DefaultHolidayLines = "2012-01-01|����" & vbCrLf & "2012-01-22|����" & vbCrLf & "2012-01-23|����" & vbCrLf & "2012-01-24|����" & vbCrLf & "2012-03-01|������" & vbCrLf & "2012-04-11|��ȸ�ǿ�������" & vbCrLf & "2012-05-05|��̳�" & vbCrLf & "2012-05-28|����ź����" & vbCrLf & "2012-06-06|������" & vbCrLf & "2012-08-15|������" & vbCrLf & "2012-09-29|�߼�" & vbCrLf & "2012-09-30|�߼�" & vbCrLf & "2012-10-01|�߼�" & vbCrLf & "2012-10-03|��õ��" & vbCrLf & "2012-12-19|����ɼ�����" & vbCrLf & "2012-12-25|�⵶ź����" & vbCrLf & ""
        Case 2013: DefaultHolidayLines = "2013-01-01|����" & vbCrLf & "2013-02-09|����" & vbCrLf & "2013-02-10|����" & vbCrLf & "2013-02-11|����" & vbCrLf & "2013-03-01|������" & vbCrLf & "2013-05-05|��� ��" & vbCrLf & "2013-05-17|����ź����" & vbCrLf & "2013-06-06|������" & vbCrLf & "2013-08-15|������" & vbCrLf & "2013-09-18|�߼�" & vbCrLf & "2013-09-19|�߼�" & vbCrLf & "2013-09-20|�߼�" & vbCrLf & "2013-10-03|��õ��" & vbCrLf & "2013-10-09|�ѱ۳�" & vbCrLf & "2013-12-25|�⵶ź����" & vbCrLf & ""
        Case 2014: DefaultHolidayLines = "2014-01-01|����" & vbCrLf & "2014-01-30|����" & vbCrLf & "2014-01-31|����" & vbCrLf & "2014-02-01|����" & vbCrLf & "2014-03-01|������" & vbCrLf & "2014-05-05|��̳�" & vbCrLf & "2014-05-06|����ź����" & vbCrLf & "2014-06-04|�������漱����" & vbCrLf & "2014-06-06|������" & vbCrLf & "2014-08-15|������" & vbCrLf & "2014-09-07|�߼�" & vbCrLf & "2014-09-08|�߼�" & vbCrLf & "2014-09-09|�߼�" & vbCrLf & "2014-09-10|��ü������" & vbCrLf & "2014-10-03|��õ��" & vbCrLf & "2014-10-09|�ѱ۳�" & vbCrLf & "2014-12-25|�⵶ź����" & vbCrLf & ""
        Case 2015: DefaultHolidayLines = "2015-01-01|����" & vbCrLf & "2015-02-18|����" & vbCrLf & "2015-02-19|����" & vbCrLf & "2015-02-20|����" & vbCrLf & "2015-03-01|������" & vbCrLf & "2015-05-05|��̳�" & vbCrLf & "2015-05-25|����ź����" & vbCrLf & "2015-06-06|������" & vbCrLf & "2015-08-15|������" & vbCrLf & "2015-09-26|�߼�" & vbCrLf & "2015-09-27|�߼�" & vbCrLf & "2015-09-28|�߼�" & vbCrLf & "2015-09-29|��ü������" & vbCrLf & "2015-10-03|��õ��" & vbCrLf & "2015-10-09|�ѱ۳�" & vbCrLf & "2015-12-25|�⵶ź����" & vbCrLf & ""
        Case 2016: DefaultHolidayLines = "2016-01-01|����" & vbCrLf & "2016-02-07|����" & vbCrLf & "2016-02-08|����" & vbCrLf & "2016-02-09|����" & vbCrLf & "2016-02-10|��ü������" & vbCrLf & "2016-03-01|������" & vbCrLf & "2016-04-13|��ȸ�ǿ�������" & vbCrLf & "2016-05-05|��̳�" & vbCrLf & "2016-05-06|�ӽð�����" & vbCrLf & "2016-05-14|����ź����" & vbCrLf & "2016-06-06|������" & vbCrLf & "2016-08-15|������" & vbCrLf & "2016-09-14|�߼�" & vbCrLf & "2016-09-15|�߼�" & vbCrLf & "2016-09-16|�߼�" & vbCrLf & "2016-10-03|��õ��" & vbCrLf & "2016-10-09|�ѱ۳�" & vbCrLf & "2016-12-25|�⵶ź����" & vbCrLf & ""
        Case 2017: DefaultHolidayLines = "2017-01-01|����" & vbCrLf & "2017-01-27|����" & vbCrLf & "2017-01-28|����" & vbCrLf & "2017-01-29|����" & vbCrLf & "2017-01-30|��ü������" & vbCrLf & "2017-03-01|������" & vbCrLf & "2017-05-03|����ź����" & vbCrLf & "2017-05-05|��̳�" & vbCrLf & "2017-05-09|����ɼ�����" & vbCrLf & "2017-06-06|������" & vbCrLf & "2017-08-15|������" & vbCrLf & "2017-10-02|�ӽð�����" & vbCrLf & "2017-10-03|�߼�" & vbCrLf & "2017-10-04|�߼�" & vbCrLf & "2017-10-05|�߼�" & vbCrLf & "2017-10-06|��ü������" & vbCrLf & "2017-10-09|�ѱ۳�" & vbCrLf & "2017-12-25|�⵶ź����" & vbCrLf & ""
        Case 2018: DefaultHolidayLines = "2018-01-01|1��1��" & vbCrLf & "2018-02-15|����" & vbCrLf & "2018-02-16|����" & vbCrLf & "2018-02-17|����" & vbCrLf & "2018-03-01|������" & vbCrLf & "2018-05-05|��̳�" & vbCrLf & "2018-05-07|��ü�޹���" & vbCrLf & "2018-05-22|��ó�Կ��ų�" & vbCrLf & "2018-06-06|������" & vbCrLf & "2018-06-13|�����������漱��" & vbCrLf & "2018-08-15|������" & vbCrLf & "2018-09-23|�߼�" & vbCrLf & "2018-09-24|�߼�" & vbCrLf & "2018-09-25|�߼�" & vbCrLf & "2018-09-26|��ü�޹���" & vbCrLf & "2018-10-03|��õ��" & vbCrLf & "2018-10-09|�ѱ۳�" & vbCrLf & "2018-12-25|�⵶ź����" & vbCrLf & ""
        Case 2019: DefaultHolidayLines = "2019-01-01|1��1��" & vbCrLf & "2019-02-04|����" & vbCrLf & "2019-02-05|����" & vbCrLf & "2019-02-06|����" & vbCrLf & "2019-03-01|������" & vbCrLf & "2019-05-05|��̳�" & vbCrLf & "2019-05-06|��ü������" & vbCrLf & "2019-05-12|��ó�Կ��ų�" & vbCrLf & "2019-06-06|������" & vbCrLf & "2019-08-15|������" & vbCrLf & "2019-09-12|�߼�" & vbCrLf & "2019-09-13|�߼�" & vbCrLf & "2019-09-14|�߼�" & vbCrLf & "2019-10-03|��õ��" & vbCrLf & "2019-10-09|�ѱ۳�" & vbCrLf & "2019-12-25|�⵶ź����" & vbCrLf & ""
        Case 2020: DefaultHolidayLines = "2020-01-01|1��1��" & vbCrLf & "2020-01-24|����" & vbCrLf & "2020-01-25|����" & vbCrLf & "2020-01-26|����" & vbCrLf & "2020-01-27|��ü������" & vbCrLf & "2020-03-01|������" & vbCrLf & "2020-04-15|��21�� ��ȸ�ǿ�����" & vbCrLf & "2020-04-30|��ó�Կ��ų�" & vbCrLf & "2020-05-05|��̳�" & vbCrLf & "2020-06-06|������" & vbCrLf & "2020-08-15|������" & vbCrLf & "2020-08-17|�ӽð�����" & vbCrLf & "2020-09-30|�߼�" & vbCrLf & "2020-10-01|�߼�" & vbCrLf & "2020-10-02|�߼�" & vbCrLf & "2020-10-03|��õ��" & vbCrLf & "2020-10-09|�ѱ۳�" & vbCrLf & "2020-12-25|�⵶ź����" & vbCrLf & ""
        Case 2021: DefaultHolidayLines = "2021-01-01|1��1��" & vbCrLf & "2021-02-11|����" & vbCrLf & "2021-02-12|����" & vbCrLf & "2021-02-13|����" & vbCrLf & "2021-03-01|������" & vbCrLf & "2021-05-05|��̳�" & vbCrLf & "2021-05-19|��ó�Կ��ų�" & vbCrLf & "2021-06-06|������" & vbCrLf & "2021-08-15|������" & vbCrLf & "2021-08-16|��ü������" & vbCrLf & "2021-09-20|�߼�" & vbCrLf & "2021-09-21|�߼�" & vbCrLf & "2021-09-22|�߼�" & vbCrLf & "2021-10-03|��õ��" & vbCrLf & "2021-10-04|��ü������" & vbCrLf & "2021-10-09|�ѱ۳�" & vbCrLf & "2021-10-11|��ü������" & vbCrLf & "2021-12-25|�⵶ź����" & vbCrLf & ""
        Case 2022: DefaultHolidayLines = "2022-01-01|1��1��" & vbCrLf & "2022-01-31|����" & vbCrLf & "2022-02-01|����" & vbCrLf & "2022-02-02|����" & vbCrLf & "2022-03-01|������" & vbCrLf & "2022-03-09|����ɼ�����" & vbCrLf & "2022-05-05|��̳�" & vbCrLf & "2022-05-08|��ó�Կ��ų�" & vbCrLf & "2022-06-01|�����������漱��" & vbCrLf & "2022-06-06|������" & vbCrLf & "2022-08-15|������" & vbCrLf & "2022-09-09|�߼�" & vbCrLf & "2022-09-10|�߼�" & vbCrLf & "2022-09-11|�߼�" & vbCrLf & "2022-09-12|��ü������" & vbCrLf & "2022-10-03|��õ��" & vbCrLf & "2022-10-09|�ѱ۳�" & vbCrLf & "2022-10-10|��ü������" & vbCrLf & "2022-12-25|�⵶ź����" & vbCrLf & ""
        Case 2023: DefaultHolidayLines = "2023-01-01|1��1��" & vbCrLf & "2023-01-21|����" & vbCrLf & "2023-01-22|����" & vbCrLf & "2023-01-23|����" & vbCrLf & "2023-01-24|��ü������" & vbCrLf & "2023-03-01|������" & vbCrLf & "2023-05-05|��̳�" & vbCrLf & "2023-05-27|��ó�Կ��ų�" & vbCrLf & "2023-05-29|��ü������" & vbCrLf & "2023-06-06|������" & vbCrLf & "2023-08-15|������" & vbCrLf & "2023-09-28|�߼�" & vbCrLf & "2023-09-29|�߼�" & vbCrLf & "2023-09-30|�߼�" & vbCrLf & "2023-10-02|�ӽð�����" & vbCrLf & "2023-10-03|��õ��" & vbCrLf & "2023-10-09|�ѱ۳�" & vbCrLf & "2023-12-25|�⵶ź����" & vbCrLf & ""
        Case 2024: DefaultHolidayLines = "2024-01-01|1��1��" & vbCrLf & "2024-02-09|����" & vbCrLf & "2024-02-10|����" & vbCrLf & "2024-02-11|����" & vbCrLf & "2024-02-12|��ü������(����)" & vbCrLf & "2024-03-01|������" & vbCrLf & "2024-04-10|��ȸ�ǿ�����" & vbCrLf & "2024-05-05|��̳�" & vbCrLf & "2024-05-06|��ü������(��̳�)" & vbCrLf & "2024-05-15|��ó�Կ��ų�" & vbCrLf & "2024-06-06|������" & vbCrLf & "2024-08-15|������" & vbCrLf & "2024-09-16|�߼�" & vbCrLf & "2024-09-17|�߼�" & vbCrLf & "2024-09-18|�߼�" & vbCrLf & "2024-10-01|�ӽð�����" & vbCrLf & "2024-10-03|��õ��" & vbCrLf & "2024-10-09|�ѱ۳�" & vbCrLf & "2024-12-25|�⵶ź����" & vbCrLf & ""
        Case 2025: DefaultHolidayLines = "2025-01-01|1��1��" & vbCrLf & "2025-01-27|�ӽð�����" & vbCrLf & "2025-01-28|����" & vbCrLf & "2025-01-29|����" & vbCrLf & "2025-01-30|����" & vbCrLf & "2025-03-01|������" & vbCrLf & "2025-03-03|��ü������" & vbCrLf & "2025-05-05|��ó�Կ��ų�" & vbCrLf & "2025-05-06|��ü������" & vbCrLf & "2025-06-03|�ӽð�����(��21�� ����� ����)" & vbCrLf & "2025-06-06|������" & vbCrLf & "2025-08-15|������" & vbCrLf & "2025-10-03|��õ��" & vbCrLf & "2025-10-05|�߼�" & vbCrLf & "2025-10-06|�߼�" & vbCrLf & "2025-10-07|�߼�" & vbCrLf & "2025-10-08|��ü������" & vbCrLf & "2025-10-09|�ѱ۳�" & vbCrLf & "2025-12-25|�⵶ź����" & vbCrLf & ""
        Case 2026: DefaultHolidayLines = "2026-01-01|1��1��" & vbCrLf & "2026-02-16|����" & vbCrLf & "2026-02-17|����" & vbCrLf & "2026-02-18|����" & vbCrLf & "2026-03-01|������" & vbCrLf & "2026-03-02|��ü������(������)" & vbCrLf & "2026-05-05|��̳�" & vbCrLf & "2026-05-24|��ó�Կ��ų�" & vbCrLf & "2026-05-25|��ü������(��ó�Կ��ų�)" & vbCrLf & "2026-06-03|�����������漱��" & vbCrLf & "2026-06-06|������" & vbCrLf & "2026-08-15|������" & vbCrLf & "2026-08-17|��ü������(������)" & vbCrLf & "2026-09-24|�߼�" & vbCrLf & "2026-09-25|�߼�" & vbCrLf & "2026-09-26|�߼�" & vbCrLf & "2026-10-03|��õ��" & vbCrLf & "2026-10-05|��ü������(��õ��)" & vbCrLf & "2026-10-09|�ѱ۳�" & vbCrLf & "2026-12-25|�⵶ź����" & vbCrLf & ""
        Case 2027: DefaultHolidayLines = "2027-01-01|1��1��" & vbCrLf & "2027-02-06|����" & vbCrLf & "2027-02-07|����" & vbCrLf & "2027-02-08|����" & vbCrLf & "2027-02-09|��ü������(����)" & vbCrLf & "2027-03-01|������" & vbCrLf & "2027-05-05|��̳�" & vbCrLf & "2027-05-13|��ó�Կ��ų�" & vbCrLf & "2027-06-06|������" & vbCrLf & "2027-08-15|������" & vbCrLf & "2027-08-16|��ü������(������)" & vbCrLf & "2027-09-14|�߼�" & vbCrLf & "2027-09-15|�߼�" & vbCrLf & "2027-09-16|�߼�" & vbCrLf & "2027-10-03|��õ��" & vbCrLf & "2027-10-04|��ü������(��õ��)" & vbCrLf & "2027-10-09|�ѱ۳�" & vbCrLf & "2027-10-11|��ü������(�ѱ۳�)" & vbCrLf & "2027-12-25|�⵶ź����" & vbCrLf & "2027-12-27|��ü������(�⵶ź����)" & vbCrLf & ""
        Case Else: DefaultHolidayLines = ""
    End Select
End Function


