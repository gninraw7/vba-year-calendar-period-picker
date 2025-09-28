VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetHolidays 
   Caption         =   "������ ���"
   ClientHeight    =   7359
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   7150
   OleObjectBlob   =   "frmSetHolidays.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmSetHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' �� 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit

'=== ���ÿ��� ���� �Ǻ� ===
Private Enum eSelOrientation
    selUnknown = 0
    selHorizontal = 1  ' 1��=����, 2��=���ϸ� (���������� ����)
    selVertical = 2    ' 1��=����, 2��=���ϸ� (��������� ����)
End Enum

Private Const REG_APP As String = "PeriodPicker"
Private Const REG_SEC As String = "Holidays"

Private Sub btnLoadFromAPI_Click()
    If txtBaseYear.text = "" Then Exit Sub
    ExportKRPublicHolidays CLng(txtBaseYear.text)
End Sub

'====== �ʱ�ȭ ======
Private Sub UserForm_Initialize()
    ' �⺻ ���� = ���翬��
    txtBaseYear.text = CStr(Year(Date))
    
    ' �� SpinButton �ʱ�ȭ
    With spnYear
        .Min = 1900
        .Max = 9999
        .SmallChange = 1
        .Value = Year(Date)
    End With
    
    ConfigListBox
    LoadYear CLng(txtBaseYear.text)
    
    btnClose.Cancel = True  ' ESC Ű�� ��� ����

End Sub

Private Sub spnYear_Change()
    On Error GoTo EH
    Dim y As Long: y = spnYear.Value
    If y < spnYear.Min Then y = spnYear.Min
    If y > spnYear.Max Then y = spnYear.Max

    txtBaseYear.text = CStr(y)
    LoadYear y
    Exit Sub
EH:
    ' ����
End Sub

Private Sub txtBaseYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim y As Long
    If Not TryGetYear(y) Then
        ' �߸� �Է� �� �ǵ���
        txtBaseYear.text = CStr(spnYear.Value)
        Exit Sub
    End If
    ' SpinButton ������ Ŭ����
    If y < spnYear.Min Then y = spnYear.Min
    If y > spnYear.Max Then y = spnYear.Max

    spnYear.Value = y
    txtBaseYear.text = CStr(y)
    LoadYear y
End Sub

Private Sub ConfigListBox()
    With lstHolidays
        .ColumnCount = 2
        .ColumnHeads = False
        .BoundColumn = 0
        .ColumnWidths = "60 pt;150 pt"  ' �ʿ信 ���� ����
        .IntegralHeight = False
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
    End With
End Sub

'====== ������Ʈ�� I/O ======
Private Function ReadHolidaysRaw(ByVal y As Long) As String
    ReadHolidaysRaw = GetSetting(REG_APP, REG_SEC, CStr(y), "")
End Function

Private Sub WriteHolidaysRaw(ByVal y As Long, ByVal raw As String)
    SaveSetting REG_APP, REG_SEC, CStr(y), raw
End Sub

'====== �ε�/���̺�(���翬�� / ������ ����) ======
Private Sub btnLoad_Click()
    Dim y As Long
    If Not TryGetYear(y) Then Exit Sub
    LoadYear y
End Sub

Private Sub LoadYear(ByVal baseYear As Long)
    lstHolidays.Clear

    Dim raw As String, lines() As String, i As Long
    raw = ReadHolidaysRaw(baseYear)
    If Len(raw) = 0 Then Exit Sub

    lines = Split(raw, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        Dim s As String: s = Trim$(lines(i))
        If Len(s) = 0 Then GoTo ContNext
        Dim a() As String: a = Split(s, "|")
        Dim ymd As String: ymd = NormalizeYMD(NzCStr(a(0)))
        If Len(ymd) = 0 Then GoTo ContNext
        Dim nm As String
        If UBound(a) >= 1 Then nm = NzCStr(a(1)) Else nm = ""
        AddListItem ymd, nm
ContNext:
    Next

    SortListBoxByDate
End Sub

Private Sub btnSaveYear_Click()
    Dim y As Long
    If Not TryGetYear(y) Then Exit Sub
    SaveYear y
    LoadYear y
End Sub

Private Sub SaveYear(ByVal baseYear As Long)
    ' ListBox �� �ش� ������ ���� �� ���� �� ����ȭ �� ������Ʈ��
    Dim keys() As String, vals() As String, n As Long
    n = ListToArrays(keys, vals)

    Dim i As Long
    Dim sb As String
    ' ���� ����
    If n > 1 Then QuickSortYMD keys, vals, 0, n - 1

    For i = 0 To n - 1
        If Left$(keys(i), 4) = CStr(baseYear) Then
            sb = sb & keys(i) & "|" & vals(i) & vbCrLf
        End If
    Next

    WriteHolidaysRaw baseYear, sb
    MsgBox baseYear & "�� ���� �Ϸ� (" & CountYearInArrays(keys, baseYear) & "��)", vbInformation
End Sub

Private Sub btnSaveAllYears_Click()
    SaveFromListSplitByYear
    Dim y As Long
    If TryGetYear(y) Then LoadYear y
End Sub

' ListBox �� (������ �׷�) �� ���� �� �� ���� Ű�� ����
Private Sub SaveFromListSplitByYear()
    Dim n As Long, i As Long
    Dim keys() As String, vals() As String
    n = ListToArrays(keys, vals)
    If n = 0 Then
        MsgBox "������ �����Ͱ� �����ϴ�.", vbInformation
        Exit Sub
    End If

    ' ������ dict
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    For i = 0 To n - 1
        Dim y As String: y = Left$(keys(i), 4)
        If Len(y) <> 4 Then GoTo ContNext
        If Not IsNumeric(y) Then GoTo ContNext
        If Not map.Exists(y) Then
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = vbTextCompare
            map.add y, d
        End If
        map(y)(keys(i)) = vals(i) ' ���� ��¥�� ������ ������ �����
ContNext:
    Next

    ' ���� Ű ����
    Dim yrs() As Variant: yrs = map.keys
    If Not IsEmpty(yrs) Then SortYearKeys yrs

    ' ������ ���ġ�����
    Dim total As Long, report As String
    For i = LBound(yrs) To UBound(yrs)
        Dim yKey As String: yKey = CStr(yrs(i))
        Dim dct As Object: Set dct = map(yKey)
        Dim dkeys() As Variant: dkeys = dct.keys
        If Not IsEmpty(dkeys) Then SortVariantYMD dkeys

        Dim sb As String, j As Long
        sb = ""
        For j = LBound(dkeys) To UBound(dkeys)
            sb = sb & CStr(dkeys(j)) & "|" & NzCStr(dct(dkeys(j))) & vbCrLf
        Next
        WriteHolidaysRaw CLng(yKey), sb
        total = total + (UBound(dkeys) - LBound(dkeys) + 1)
        report = report & vbCrLf & yKey & "��: " & (UBound(dkeys) - LBound(dkeys) + 1) & "��"
    Next

    MsgBox "������ ���� ���� �Ϸ� (" & total & "��):" & report, vbInformation
End Sub

'====== ���ÿ������� �ҷ�����(����/���� �ڵ�+����) ======
Private Sub btnLoadFromSelection_Click()
    On Error GoTo EH

    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "Ȱ�� ��Ʈ���� ������ ������ �� �����ϼ���.", vbExclamation
        Exit Sub
    End If
    If rng.Rows.Count < 2 And rng.Columns.Count < 2 Then
        MsgBox "���ÿ����� �ּ� 2�� �Ǵ� 2���̾�� �մϴ�." & vbCrLf & _
               "(������: 1��=����, 2��=���ϸ� / ������: 1��=����, 2��=���ϸ�)", vbExclamation
        Exit Sub
    End If

    Dim orient As eSelOrientation
    orient = ResolveSelectionOrientation(rng)

    If orient = selUnknown Then
        MsgBox "���ÿ����� ������/������ ��� ������ �Ǵ��� �� �����ϴ�.", vbExclamation
        Exit Sub
    End If

    ' Dict�� ����(�ߺ� ��¥�� ������ ������)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long, c As Long
    If orient = selHorizontal Then
        For c = 1 To rng.Columns.Count
            Dim rawD As Variant, rawN As Variant
            rawD = rng.Cells(1, c).Value
            rawN = rng.Cells(2, c).Value
    
            ' �� ���� ��ŵ
            If Len(Trim$(NzCStr(rawD))) = 0 Then GoTo ContinueC
    
            ' ���� ����ȭ ����(���� �ƴ�) ��ŵ
            Dim ymdH As String
            ymdH = CanonYMDFromCell(rawD)
            If Len(ymdH) = 0 Then GoTo ContinueC
    
            dict(ymdH) = NzCStr(rawN)   ' �̸��� ���� ����
ContinueC:
        Next
    Else
        For r = 1 To rng.Rows.Count
            Dim rawDv As Variant, rawNv As Variant
            rawDv = rng.Cells(r, 1).Value
            rawNv = rng.Cells(r, 2).Value
    
            ' �� ���� ��ŵ
            If Len(Trim$(NzCStr(rawDv))) = 0 Then GoTo ContinueR
    
            ' ���� ����ȭ ����(���� �ƴ�) ��ŵ
            Dim ymdV As String
            ymdV = CanonYMDFromCell(rawDv)
            If Len(ymdV) = 0 Then GoTo ContinueR
    
            dict(ymdV) = NzCStr(rawNv)
ContinueR:
        Next
    End If

    If dict.Count = 0 Then
        MsgBox "��ȿ�� ��¥�� ã�� ���߽��ϴ�. (yyyy-mm-dd �Ǵ� ��ȿ�� Excel ��¥)", vbExclamation
        Exit Sub
    End If

    ' (�ɼ�) ȥ�� ���� ���
    If chkWarnMixedYears.Value Then
        If Not ConfirmMixedYears_OnDict(dict, "�ҷ��� �����Ϳ� ���� �ٸ� ������ ���� �ֽ��ϴ�. �׷��� �ݿ��ұ��?") Then Exit Sub
    End If

    ' ListBox�� �ݿ�(�����Ͽ� ǥ��)
    Dim keys As Variant: keys = dict.keys
    SortVariantYMD keys

    lstHolidays.Clear
    Dim ix
    For Each ix In keys
        AddListItem CStr(ix), NzCStr(dict(ix))
    Next

    ' ���س⵵ �ڵ� ����(��� ������ ù �׸� ����)
    If lstHolidays.ListCount > 0 Then
        If Len(Trim$(txtBaseYear.text)) = 0 Then
            txtBaseYear.text = Left$(CStr(lstHolidays.list(0, 0)), 4)
        End If
    End If

    MsgBox "���ÿ������� " & lstHolidays.ListCount & "���� �ҷ��Խ��ϴ�." & vbCrLf & _
           "����: " & IIf(orient = selHorizontal, "������", "������"), vbInformation
    Exit Sub
EH:
    MsgBox "���ÿ��� �ҷ����� �� ����: " & Err.Description, vbExclamation
End Sub

'====== ListBox �߰�/����/����/����Ŭ�� ���� ======
Private Sub btnAddUpdate_Click()
    Dim ymd As String: ymd = NormalizeYMD(Trim$(txtHoliday.text))
    If Len(ymd) = 0 Then
        MsgBox "������ yyyy-mm-dd �������� �Է��ϼ���.", vbExclamation
        txtHoliday.SetFocus
        Exit Sub
    End If
    Dim nm As String: nm = Trim$(txtName.text)

    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx >= 0 Then
        ' ����
        lstHolidays.list(idx, 0) = ymd
        lstHolidays.list(idx, 1) = nm
    Else
        ' �߰�
        AddListItem ymd, nm
    End If

    SortListBoxByDate
    ClearEditBoxes
End Sub

Private Sub btnDelete_Click()
    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx < 0 Then
        MsgBox "������ �׸��� �����ϼ���.", vbExclamation
        Exit Sub
    End If
    lstHolidays.RemoveItem idx
    ClearEditBoxes
End Sub

Private Sub lstHolidays_Click()
    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx >= 0 Then
        txtHoliday.text = NzCStr(lstHolidays.list(idx, 0))
        txtName.text = NzCStr(lstHolidays.list(idx, 1))
    End If
End Sub

Private Sub lstHolidays_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx < 0 Then Exit Sub

    Dim cur As String
    cur = NzCStr(lstHolidays.list(idx, 0)) & "," & NzCStr(lstHolidays.list(idx, 1))

    Dim s As String
    s = InputBox("����,���ϸ� �������� �Է��ϼ���" & vbLf & " (��: 2025-01-01,����)", "�׸� ����", cur)
    If Len(Trim$(s)) = 0 Then Exit Sub

    Dim ymd As String, nm As String
    If Not TryParseYMDName(s, ymd, nm) Then
        MsgBox "�Է� ������ �ùٸ��� �ʽ��ϴ�. ��) 2025-01-01,����", vbExclamation
        Exit Sub
    End If

    lstHolidays.list(idx, 0) = ymd
    lstHolidays.list(idx, 1) = nm
    SortListBoxByDate
End Sub

Private Sub btnExportSheet_Click()
    On Error GoTo EH
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
    On Error Resume Next
    ws.name = "Holidays_" & Format(Now, "yymmdd_hhnnss")
    On Error GoTo 0

    ws.Range("A1").Value = "����"
    ws.Range("B1").Value = "���ϸ�"
    ws.Range("A1:B1").Font.Bold = True

    Dim i As Long
    For i = 0 To lstHolidays.ListCount - 1
        ws.Cells(i + 2, 1).Value = CDate(lstHolidays.list(i, 0))
        ws.Cells(i + 2, 2).Value = NzCStr(lstHolidays.list(i, 1))
    Next

    ws.Columns(1).NumberFormat = "yyyy-mm-dd"
    ws.Columns("A:B").AutoFit

    MsgBox "���� ����� �� ��Ʈ '" & ws.name & "' �� ����߽��ϴ�. (" & lstHolidays.ListCount & "��)", vbInformation
    Exit Sub
EH:
    MsgBox "Sheet ��� �� ����: " & Err.Description, vbExclamation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'====== ��ƿ: ListBox ����/����/��ȯ ======
Private Sub AddListItem(ByVal ymd As String, ByVal nm As String)
    Dim ix As Long
    lstHolidays.AddItem ymd
    ix = lstHolidays.ListCount - 1
    lstHolidays.list(ix, 1) = nm
End Sub

Private Sub ClearEditBoxes()
    txtHoliday.text = ""
    txtName.text = ""
    lstHolidays.ListIndex = -1
End Sub

Private Function ListToArrays(ByRef keys() As String, ByRef vals() As String) As Long
    Dim n As Long: n = lstHolidays.ListCount
    If n <= 0 Then ListToArrays = 0: Exit Function
    ReDim keys(0 To n - 1)
    ReDim vals(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        keys(i) = NormalizeYMD(NzCStr(lstHolidays.list(i, 0)))
        vals(i) = NzCStr(lstHolidays.list(i, 1))
    Next
    ListToArrays = n
End Function

Private Sub SortListBoxByDate()
    Dim keys() As String, vals() As String, n As Long
    n = ListToArrays(keys, vals)
    If n <= 1 Then Exit Sub
    QuickSortYMD keys, vals, 0, n - 1

    ' ����ε�
    lstHolidays.Clear
    Dim i As Long
    For i = 0 To n - 1
        AddListItem keys(i), vals(i)
    Next
End Sub

Private Sub QuickSortYMD(ByRef a() As String, ByRef b() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String
    i = lo: j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            Dim t As String
            t = a(i): a(i) = a(j): a(j) = t
            t = b(i): b(i) = b(j): b(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortYMD a, b, lo, j
    If i < hi Then QuickSortYMD a, b, i, hi
End Sub

'Private Sub SortVariantYMD(ByRef arr() As Variant)
Private Sub SortVariantYMD(ByRef arr As Variant)
    Dim i As Long, j As Long, t As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(i)) > CStr(arr(j)) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next
    Next
End Sub

Private Sub SortYearKeys(ByRef yrs() As Variant)
    Dim i As Long, j As Long, t As Variant
    For i = LBound(yrs) To UBound(yrs) - 1
        For j = i + 1 To UBound(yrs)
            If CLng(yrs(i)) > CLng(yrs(j)) Then t = yrs(i): yrs(i) = yrs(j): yrs(j) = t
        Next
    Next
End Sub

Private Function CountYearInArrays(ByRef keys() As String, ByVal y As Long) As Long
    Dim i As Long, cnt As Long
    For i = LBound(keys) To UBound(keys)
        If Left$(keys(i), 4) = CStr(y) Then cnt = cnt + 1
    Next
    CountYearInArrays = cnt
End Function

'====== ���ÿ��� ����/���� ����/�ڵ� �Ǻ� ======
Private Function ResolveSelectionOrientation(ByVal rng As Range) As eSelOrientation
    If chkForceHorizontal.Value And Not chkForceVertical.Value Then
        ResolveSelectionOrientation = selHorizontal: Exit Function
    ElseIf chkForceVertical.Value And Not chkForceHorizontal.Value Then
        ResolveSelectionOrientation = selVertical: Exit Function
    ElseIf chkForceHorizontal.Value And chkForceVertical.Value Then
        ResolveSelectionOrientation = selVertical: Exit Function
    End If
    ResolveSelectionOrientation = DetectSelectionOrientation(rng)
End Function

Private Function DetectSelectionOrientation(ByVal rng As Range) As eSelOrientation
    On Error Resume Next
    Dim scoreH As Long, scoreV As Long
    If rng.Rows.Count >= 1 Then scoreH = CountValidDatesInRow(rng, 1)
    If rng.Columns.Count >= 1 Then scoreV = CountValidDatesInCol(rng, 1)

    Dim canH As Boolean: canH = (rng.Rows.Count >= 2)
    Dim canV As Boolean: canV = (rng.Columns.Count >= 2)

    If Not canH And Not canV Then DetectSelectionOrientation = selUnknown: Exit Function

    If canH And canV Then
        If scoreH > scoreV Then
            DetectSelectionOrientation = selHorizontal
        ElseIf scoreV > scoreH Then
            DetectSelectionOrientation = selVertical
        Else
            If rng.Rows.Count = 2 And rng.Columns.Count > 2 Then
                DetectSelectionOrientation = selHorizontal
            ElseIf rng.Columns.Count = 2 And rng.Rows.Count > 2 Then
                DetectSelectionOrientation = selVertical
            Else
                DetectSelectionOrientation = selUnknown
            End If
        End If
    ElseIf canH Then
        DetectSelectionOrientation = selHorizontal
    Else
        DetectSelectionOrientation = selVertical
    End If
End Function

Private Function CountValidDatesInRow(ByVal rng As Range, ByVal rowIdx As Long) As Long
    Dim c As Long, cnt As Long
    For c = 1 To rng.Columns.Count
        If Len(CanonYMDFromCell(rng.Cells(rowIdx, c).Value)) > 0 Then cnt = cnt + 1
    Next
    CountValidDatesInRow = cnt
End Function

Private Function CountValidDatesInCol(ByVal rng As Range, ByVal colIdx As Long) As Long
    Dim r As Long, cnt As Long
    For r = 1 To rng.Rows.Count
        If Len(CanonYMDFromCell(rng.Cells(r, colIdx).Value)) > 0 Then cnt = cnt + 1
    Next
    CountValidDatesInCol = cnt
End Function

Private Function ConfirmMixedYears_OnDict(ByVal dict As Object, ByVal prompt As String) As Boolean
    Dim years As Object: Set years = CreateObject("Scripting.Dictionary")
    years.CompareMode = vbTextCompare
    Dim k As Variant
    For Each k In dict.keys
        If Len(k) >= 4 Then years(Left$(CStr(k), 4)) = True
    Next
    If years.Count <= 1 Then ConfirmMixedYears_OnDict = True: Exit Function

    Dim list As String: list = ""
    For Each k In years.keys
        If Len(list) > 0 Then list = list & ", "
        list = list & CStr(k)
    Next
    ConfirmMixedYears_OnDict = (MsgBox(prompt & vbCrLf & "(����: " & list & ")", _
                              vbExclamation + vbYesNo, "ȥ�� ���� ����") = vbYes)
End Function


Private Function TryGetYear(ByRef y As Long) As Boolean
    On Error GoTo EH
    y = CLng(Trim$(txtBaseYear.text))
    If y < 1900 Or y > 9999 Then
        MsgBox "���س⵵�� ��Ȯ�� �Է��ϼ���. ��) 2025", vbExclamation
        Exit Function
    End If
    TryGetYear = True
    Exit Function
EH:
    MsgBox "���س⵵�� ��Ȯ�� �Է��ϼ���. ��) 2025", vbExclamation
End Function

