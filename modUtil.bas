Attribute VB_Name = "modUtil"
Option Explicit

' "yyyy-mm-dd,이름" → ymd/name 로 분해, ymd 정규화
Public Function TryParseYMDName(ByVal s As String, ByRef ymd As String, ByRef nm As String) As Boolean
    s = Trim$(s)
    If s = "" Then Exit Function
    s = Replace(s, "，", ",") ' 전각 콤마 허용

    Dim p As Long: p = InStr(1, s, ",")
    If p = 0 Then Exit Function

    Dim lefts As String, rights As String
    lefts = Trim$(Left$(s, p - 1))
    rights = Trim$(Mid$(s, p + 1))

    ymd = NormalizeYMD(lefts)  ' 기존 정규화 함수 재사용
    If Len(ymd) = 0 Then Exit Function

    nm = rights
    TryParseYMDName = True
End Function

' ==== TryParse: "yyyy-mm-dd" → Date ====
Public Function TryParseYMD(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    s = Trim$(s)
    If s Like "####-##-##" Then
        d = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Right$(s, 2)))
        TryParseYMD = True
    End If
    Exit Function
EH:
    TryParseYMD = False
End Function

' ==== 주말 여부 ====
Public Function IsWeekend(ByVal d As Date) As Boolean
    Dim wd As VbDayOfWeek: wd = Weekday(d, vbSunday)
    IsWeekend = (wd = vbSunday Or wd = vbSaturday)
End Function

' ================== "yyyy-mm-dd" → Date 변환 (구분자 혼합/오염 방지) ==================
Public Function TryParseYMDToDate(ByVal s As String, ByRef dtOut As Date) As Boolean
    On Error GoTo EH
    Dim a() As String, y As Long, m As Long, d As Long
    s = Trim$(s)
    If s = "" Then Exit Function

    ' 구분자 정규화
    s = Replace(s, ".", "-")
    s = Replace(s, "/", "-")

    a = Split(s, "-")
    If UBound(a) <> 2 Then Exit Function
    y = CLng(Trim$(a(0))): m = CLng(Trim$(a(1))): d = CLng(Trim$(a(2)))

    dtOut = DateSerial(y, m, d)   ' 유효성 자동 검증됨
    TryParseYMDToDate = True
    Exit Function
EH:
    TryParseYMDToDate = False
End Function

Public Function NzCStr(ByVal v As Variant) As String
    If IsError(v) Or isNull(v) Or IsEmpty(v) Then NzCStr = "" Else NzCStr = CStr(v)
End Function

Public Function Nz(ByVal v, Optional ByVal def As Variant = "") As Variant
    If IsEmpty(v) Or isNull(v) Then
        Nz = def
    Else
        Nz = v
    End If
End Function

Public Function TryParseDate(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    If Len(Trim$(s)) = 0 Then
        TryParseDate = False
        Exit Function
    End If
    d = CDate(s)
    TryParseDate = True
    Exit Function
EH:
    TryParseDate = False
End Function


'====== 유틸: 파서/검증 ======
Public Function CanonYMDFromCell(ByVal v As Variant) As String
    On Error GoTo Fallback
    If IsDate(v) Then
        CanonYMDFromCell = Format$(CDate(v), "yyyy-mm-dd")
        Exit Function
    End If
Fallback:
    On Error Resume Next
    Dim s As String: s = Trim$(NzCStr(v))
    If s = "" Then Exit Function
    s = Replace(s, ".", "-")
    s = Replace(s, "/", "-")
    CanonYMDFromCell = NormalizeYMD(s)
End Function

Public Function NormalizeYMD(ByVal s As String) As String
    On Error GoTo EH
    Dim a() As String, dT As Date
    s = Trim$(s)
    If s = "" Then Exit Function
    s = Replace(s, ".", "-")
    s = Replace(s, "/", "-")
    a = Split(s, "-")
    If UBound(a) <> 2 Then Exit Function
    dT = DateSerial(CLng(a(0)), CLng(a(1)), CLng(a(2)))
    NormalizeYMD = Format$(dT, "yyyy-mm-dd")
    Exit Function
EH:
    NormalizeYMD = ""
End Function

Public Function SanitizeFileName(ByVal s As String) As String
    Dim bad: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    s = Trim$(s)
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next
    ' 끝의 점/공백 제거
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    If s = "" Then s = "tasks"
    SanitizeFileName = s
End Function

Public Function EscapeLikePattern(ByVal p As String) As String
    ' [, ], *, ?, #, \  → \로 이스케이프
    Dim ch As String, i As Long, out As String
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If ch Like "[\[\]\*\?#\]" Then out = out & "\" & ch Else out = out & ch
    Next
    EscapeLikePattern = out
End Function

