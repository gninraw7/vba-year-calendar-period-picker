VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetHolidays 
   Caption         =   "공휴일 등록"
   ClientHeight    =   7359
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   7150
   OleObjectBlob   =   "frmSetHolidays.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmSetHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit

'=== 선택영역 형태 판별 ===
Private Enum eSelOrientation
    selUnknown = 0
    selHorizontal = 1  ' 1행=휴일, 2행=휴일명 (열방향으로 진행)
    selVertical = 2    ' 1열=휴일, 2열=휴일명 (행방향으로 진행)
End Enum

Private Const REG_APP As String = "PeriodPicker"
Private Const REG_SEC As String = "Holidays"

Private Sub btnLoadFromAPI_Click()
    If txtBaseYear.text = "" Then Exit Sub
    ExportKRPublicHolidays CLng(txtBaseYear.text)
End Sub

'====== 초기화 ======
Private Sub UserForm_Initialize()
    ' 기본 연도 = 현재연도
    txtBaseYear.text = CStr(Year(Date))
    
    ' ▼ SpinButton 초기화
    With spnYear
        .Min = 1900
        .Max = 9999
        .SmallChange = 1
        .Value = Year(Date)
    End With
    
    ConfigListBox
    LoadYear CLng(txtBaseYear.text)
    
    btnClose.Cancel = True  ' ESC 키로 취소 가능

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
    ' 무시
End Sub

Private Sub txtBaseYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim y As Long
    If Not TryGetYear(y) Then
        ' 잘못 입력 시 되돌림
        txtBaseYear.text = CStr(spnYear.Value)
        Exit Sub
    End If
    ' SpinButton 범위로 클램프
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
        .ColumnWidths = "60 pt;150 pt"  ' 필요에 따라 조정
        .IntegralHeight = False
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
    End With
End Sub

'====== 레지스트리 I/O ======
Private Function ReadHolidaysRaw(ByVal y As Long) As String
    ReadHolidaysRaw = GetSetting(REG_APP, REG_SEC, CStr(y), "")
End Function

Private Sub WriteHolidaysRaw(ByVal y As Long, ByVal raw As String)
    SaveSetting REG_APP, REG_SEC, CStr(y), raw
End Sub

'====== 로드/세이브(현재연도 / 연도별 분할) ======
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
    ' ListBox → 해당 연도만 추출 → 정렬 → 직렬화 → 레지스트리
    Dim keys() As String, vals() As String, n As Long
    n = ListToArrays(keys, vals)

    Dim i As Long
    Dim sb As String
    ' 먼저 정렬
    If n > 1 Then QuickSortYMD keys, vals, 0, n - 1

    For i = 0 To n - 1
        If Left$(keys(i), 4) = CStr(baseYear) Then
            sb = sb & keys(i) & "|" & vals(i) & vbCrLf
        End If
    Next

    WriteHolidaysRaw baseYear, sb
    MsgBox baseYear & "년 저장 완료 (" & CountYearInArrays(keys, baseYear) & "건)", vbInformation
End Sub

Private Sub btnSaveAllYears_Click()
    SaveFromListSplitByYear
    Dim y As Long
    If TryGetYear(y) Then LoadYear y
End Sub

' ListBox → (연도별 그룹) → 정렬 → 각 연도 키로 저장
Private Sub SaveFromListSplitByYear()
    Dim n As Long, i As Long
    Dim keys() As String, vals() As String
    n = ListToArrays(keys, vals)
    If n = 0 Then
        MsgBox "저장할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' 연도별 dict
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
        map(y)(keys(i)) = vals(i) ' 같은 날짜는 마지막 값으로 덮어쓰기
ContNext:
    Next

    ' 연도 키 정렬
    Dim yrs() As Variant: yrs = map.keys
    If Not IsEmpty(yrs) Then SortYearKeys yrs

    ' 연도별 정렬→저장
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
        report = report & vbCrLf & yKey & "년: " & (UBound(dkeys) - LBound(dkeys) + 1) & "건"
    Next

    MsgBox "연도별 분할 저장 완료 (" & total & "건):" & report, vbInformation
End Sub

'====== 선택영역에서 불러오기(가로/세로 자동+강제) ======
Private Sub btnLoadFromSelection_Click()
    On Error GoTo EH

    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "활성 시트에서 범위를 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    If rng.Rows.Count < 2 And rng.Columns.Count < 2 Then
        MsgBox "선택영역은 최소 2행 또는 2열이어야 합니다." & vbCrLf & _
               "(가로형: 1행=휴일, 2행=휴일명 / 세로형: 1열=휴일, 2열=휴일명)", vbExclamation
        Exit Sub
    End If

    Dim orient As eSelOrientation
    orient = ResolveSelectionOrientation(rng)

    If orient = selUnknown Then
        MsgBox "선택영역이 가로형/세로형 어느 쪽인지 판단할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' Dict로 수집(중복 날짜는 마지막 값으로)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long, c As Long
    If orient = selHorizontal Then
        For c = 1 To rng.Columns.Count
            Dim rawD As Variant, rawN As Variant
            rawD = rng.Cells(1, c).Value
            rawN = rng.Cells(2, c).Value
    
            ' 값 없음 스킵
            If Len(Trim$(NzCStr(rawD))) = 0 Then GoTo ContinueC
    
            ' 일자 정규화 실패(형식 아님) 스킵
            Dim ymdH As String
            ymdH = CanonYMDFromCell(rawD)
            If Len(ymdH) = 0 Then GoTo ContinueC
    
            dict(ymdH) = NzCStr(rawN)   ' 이름은 공백 가능
ContinueC:
        Next
    Else
        For r = 1 To rng.Rows.Count
            Dim rawDv As Variant, rawNv As Variant
            rawDv = rng.Cells(r, 1).Value
            rawNv = rng.Cells(r, 2).Value
    
            ' 값 없음 스킵
            If Len(Trim$(NzCStr(rawDv))) = 0 Then GoTo ContinueR
    
            ' 일자 정규화 실패(형식 아님) 스킵
            Dim ymdV As String
            ymdV = CanonYMDFromCell(rawDv)
            If Len(ymdV) = 0 Then GoTo ContinueR
    
            dict(ymdV) = NzCStr(rawNv)
ContinueR:
        Next
    End If

    If dict.Count = 0 Then
        MsgBox "유효한 날짜를 찾지 못했습니다. (yyyy-mm-dd 또는 유효한 Excel 날짜)", vbExclamation
        Exit Sub
    End If

    ' (옵션) 혼합 연도 경고
    If chkWarnMixedYears.Value Then
        If Not ConfirmMixedYears_OnDict(dict, "불러온 데이터에 서로 다른 연도가 섞여 있습니다. 그래도 반영할까요?") Then Exit Sub
    End If

    ' ListBox에 반영(정렬하여 표시)
    Dim keys As Variant: keys = dict.keys
    SortVariantYMD keys

    lstHolidays.Clear
    Dim ix
    For Each ix In keys
        AddListItem CStr(ix), NzCStr(dict(ix))
    Next

    ' 기준년도 자동 설정(비어 있으면 첫 항목 연도)
    If lstHolidays.ListCount > 0 Then
        If Len(Trim$(txtBaseYear.text)) = 0 Then
            txtBaseYear.text = Left$(CStr(lstHolidays.list(0, 0)), 4)
        End If
    End If

    MsgBox "선택영역에서 " & lstHolidays.ListCount & "건을 불러왔습니다." & vbCrLf & _
           "형태: " & IIf(orient = selHorizontal, "가로형", "세로형"), vbInformation
    Exit Sub
EH:
    MsgBox "선택영역 불러오기 중 오류: " & Err.Description, vbExclamation
End Sub

'====== ListBox 추가/수정/삭제/더블클릭 수정 ======
Private Sub btnAddUpdate_Click()
    Dim ymd As String: ymd = NormalizeYMD(Trim$(txtHoliday.text))
    If Len(ymd) = 0 Then
        MsgBox "휴일을 yyyy-mm-dd 형식으로 입력하세요.", vbExclamation
        txtHoliday.SetFocus
        Exit Sub
    End If
    Dim nm As String: nm = Trim$(txtName.text)

    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx >= 0 Then
        ' 수정
        lstHolidays.list(idx, 0) = ymd
        lstHolidays.list(idx, 1) = nm
    Else
        ' 추가
        AddListItem ymd, nm
    End If

    SortListBoxByDate
    ClearEditBoxes
End Sub

Private Sub btnDelete_Click()
    Dim idx As Long: idx = lstHolidays.ListIndex
    If idx < 0 Then
        MsgBox "삭제할 항목을 선택하세요.", vbExclamation
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
    s = InputBox("휴일,휴일명 형식으로 입력하세요" & vbLf & " (예: 2025-01-01,신정)", "항목 수정", cur)
    If Len(Trim$(s)) = 0 Then Exit Sub

    Dim ymd As String, nm As String
    If Not TryParseYMDName(s, ymd, nm) Then
        MsgBox "입력 형식이 올바르지 않습니다. 예) 2025-01-01,신정", vbExclamation
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

    ws.Range("A1").Value = "휴일"
    ws.Range("B1").Value = "휴일명"
    ws.Range("A1:B1").Font.Bold = True

    Dim i As Long
    For i = 0 To lstHolidays.ListCount - 1
        ws.Cells(i + 2, 1).Value = CDate(lstHolidays.list(i, 0))
        ws.Cells(i + 2, 2).Value = NzCStr(lstHolidays.list(i, 1))
    Next

    ws.Columns(1).NumberFormat = "yyyy-mm-dd"
    ws.Columns("A:B").AutoFit

    MsgBox "현재 목록을 새 시트 '" & ws.name & "' 에 출력했습니다. (" & lstHolidays.ListCount & "건)", vbInformation
    Exit Sub
EH:
    MsgBox "Sheet 출력 중 오류: " & Err.Description, vbExclamation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'====== 유틸: ListBox 조작/정렬/변환 ======
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

    ' 재바인딩
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

'====== 선택영역 가로/세로 강제/자동 판별 ======
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
    ConfirmMixedYears_OnDict = (MsgBox(prompt & vbCrLf & "(연도: " & list & ")", _
                              vbExclamation + vbYesNo, "혼합 연도 감지") = vbYes)
End Function


Private Function TryGetYear(ByRef y As Long) As Boolean
    On Error GoTo EH
    y = CLng(Trim$(txtBaseYear.text))
    If y < 1900 Or y > 9999 Then
        MsgBox "기준년도를 정확히 입력하세요. 예) 2025", vbExclamation
        Exit Function
    End If
    TryGetYear = True
    Exit Function
EH:
    MsgBox "기준년도를 정확히 입력하세요. 예) 2025", vbExclamation
End Function

