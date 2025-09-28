Attribute VB_Name = "modTaskStoreJSON"
'==== modTaskStoreJSON (추가/교체) =================================
Option Explicit

' 병합 사용 여부(안전 기본값: False)
Private Const MERGE_ADJACENT_ENABLED As Boolean = False

Private Sub InitJsonOptions()
    With JsonConverter.JsonOptions
        .UseDoubleForLargeNumbers = True
        .AllowUnquotedKeys = False
        .EscapeSolidus = False
    End With
End Sub

Private Function FixTrailingCommas(ByVal s As String) As String
    ' 매우 단순한 보정: ",]" -> "]", ",}" -> "}"
    FixTrailingCommas = Replace(Replace(s, ",]", "]"), ",}", "}")
End Function

Private Function GetJsonStr(ByVal d As Object, ByVal k As String) As String
    On Error Resume Next
    If Not d Is Nothing Then
        If d.Exists(k) Then GetJsonStr = CStr(d(k))
    End If
End Function

' UTF-8 BOM 제거 +(선택) 간단 전처리
Private Function NormalizeJson(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2) ' UTF-8 BOM 제거
    End If
    ' 주석/트레일링 콤마 허용 안 하므로, 저장 쪽에서 올바른 JSON을 만들도록 유지하세요.
    NormalizeJson = s
End Function

' Dictionary/Collection을 안전하게 clsTaskItem으로 변환
Private Sub AddTaskFromDict(ByRef out As Collection, ByVal d As Object)
    Dim t As New clsTaskItem
    Dim v As Variant, dF As Date, dT As Date, okF As Boolean, okT As Boolean

    On Error GoTo EH

    v = GetJsonValue(d, "name"): t.TaskName = NzCStr(v)

    v = GetJsonValue(d, "from")
    okF = (TypeName(v) = "Date")
    If Not okF Then okF = TryParseYMD(NzCStr(v), dF) Or TryParseDate(NzCStr(v), dF)
    If Not okF Then Exit Sub
    If TypeName(v) = "Date" Then dF = v
    t.FromDate = dF

    v = GetJsonValue(d, "to")
    okT = (Not IsEmpty(v) And Not isNull(v) And Len(NzCStr(v)) > 0)
    If okT Then
        If TypeName(v) = "Date" Then
            dT = v
        ElseIf Not (TryParseYMD(NzCStr(v), dT) Or TryParseDate(NzCStr(v), dT)) Then
            okT = False
        End If
    End If
    t.HasTo = okT
    If okT Then t.ToDate = dT

    out.add t
    Exit Sub
EH:
    ' 한 항목이 틀려도 전체는 계속
End Sub

' Dictionary에서 키 안전 추출
Private Function GetJsonValue(ByVal d As Object, ByVal key As String) As Variant
    On Error GoTo EH
    If TypeName(d) = "Dictionary" Then
        If d.Exists(key) Then GetJsonValue = d(key) Else GetJsonValue = Empty
    Else
        GetJsonValue = Empty
    End If
    Exit Function
EH:
    GetJsonValue = Empty
End Function

' === 최종: 텍스트 → Collection(clsTaskItem) ===
' modTaskStoreJSON
Private Function JsonTextToTasks(ByVal json As String) As Collection
    On Error GoTo EH

    ' 필요 시 BOM/말줄임 쉼표 정리(이미 있으시면 그대로 두세요)
    json = NormalizeJson(json)

    ' <<< 중요: 이 버전은 점(.) 속성으로만 설정 가능 >>>
    With JsonConverter.JsonOptions
        .UseDoubleForLargeNumbers = True   ' 또는 False, 프로젝트 정책에 맞게
        .AllowUnquotedKeys = False
        .EscapeSolidus = False
    End With

    Dim root As Variant
    Set root = JsonConverter.ParseJson(json)  ' 인수 1개만!

    Dim out As New Collection
    Dim i As Long

    If TypeName(root) = "Collection" Then
        For i = 1 To root.Count
            Dim t As clsTaskItem
            Set t = TaskFromDict(root(i))
            If Not t Is Nothing Then out.add t
        Next
    ElseIf TypeName(root) = "Dictionary" Then
        ' 혹시 {"tasks":[...]} 구조도 허용하려면
        If root.Exists("tasks") Then
            Dim arr As Collection
            Set arr = root("tasks")
            For i = 1 To arr.Count
                Dim t2 As clsTaskItem
                Set t2 = TaskFromDict(arr(i))
                If Not t2 Is Nothing Then out.add t2
            Next
        End If
    End If

    Set JsonTextToTasks = out
    Exit Function
EH:
    MsgBox "JSON 파싱 오류 (" & Err.Number & "): " & Err.Description, vbExclamation
    Dim z As New Collection
    Set JsonTextToTasks = z
End Function

Private Function TaskFromDict(ByVal d As Object) As clsTaskItem
    On Error GoTo EH
    If TypeName(d) <> "Dictionary" Then Exit Function

    Dim nm As String, sf As String, st As String
    If d.Exists("name") Then nm = CStr(d("name"))
    If d.Exists("from") Then sf = CStr(d("from"))
    If d.Exists("to") Then st = IIf(IsEmpty(d("to")) Or isNull(d("to")), "", CStr(d("to")))

    If Len(sf) = 0 Then Exit Function

    Dim dF As Date, dT As Date, okF As Boolean, okT As Boolean
    okF = TryParseYMD(sf, dF) Or TryParseDate(sf, dF)
    If Not okF Then Exit Function

    Dim t As New clsTaskItem
    t.TaskName = nm
    t.FromDate = dF

    okT = (Len(st) > 0) And (TryParseYMD(st, dT) Or TryParseDate(st, dT))
    t.HasTo = okT
    If okT Then t.ToDate = dT

    Set TaskFromDict = t
    Exit Function
EH:
    Set TaskFromDict = Nothing
End Function

' 단일 파일: %APPDATA%\PeriodPicker\Tasks\tasks.json
Public Function TaskDataRootFolder() As String
    Dim p As String: p = Environ$("APPDATA") & "\PeriodPicker\Tasks"
    EnsureFolder p
    TaskDataRootFolder = p
End Function

Public Function AllTasksJsonPath() As String
    AllTasksJsonPath = TaskDataRootFolder() & "\tasks.json"
End Function

Public Sub EnsureFolder(ByVal p As String)
    Dim fso As Object, parts As Variant, cur As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(p) Then Exit Sub
    parts = Split(p, "\"): cur = parts(0)
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If Not fso.FolderExists(cur) Then fso.CreateFolder cur
        End If
    Next
End Sub

Private Function ParentFolderOf(ByVal path As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ParentFolderOf = fso.GetParentFolderName(path)
End Function

'--- 파일 I/O (UTF-8, 안전 저장) ---
Private Function ReadAllTextUTF8(ByVal path As String) As String
    On Error GoTo EH
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2: .Charset = "utf-8": .Open
        .LoadFromFile path
        ReadAllTextUTF8 = .ReadText(-1)
        .Close
    End With
    Exit Function
EH:
    ReadAllTextUTF8 = ""
End Function

Private Sub SafeWriteAllTextUTF8(ByVal path As String, ByVal text As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder ParentFolderOf(path)
    Dim tmp As String: tmp = path & ".tmp"

    On Error Resume Next
    If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    If fso.FileExists(path) Then
        If (GetAttr(path) And vbReadOnly) <> 0 Then SetAttr path, (GetAttr(path) And Not vbReadOnly)
    End If
    On Error GoTo 0

    Dim ts As Object, bs As Object
    Set ts = CreateObject("ADODB.Stream")
    Set bs = CreateObject("ADODB.Stream")

    On Error GoTo Fallback

    ' 1) Text 스트림에 UTF-8로 기록
    With ts
        .Type = 2              ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText text, 0     ' adWriteLine=1 아님
        .Position = 0
    End With

    ' 2) Binary 스트림으로 복사 후 저장
    With bs
        .Type = 1              ' adTypeBinary
        .Open
        ts.CopyTo bs
        ts.Close               ' 소스 먼저 닫아 잠금 해제
        .Position = 0
        .SaveToFile tmp, 2     ' adSaveCreateOverWrite
        .Close
    End With

    On Error Resume Next
    If fso.FileExists(path) Then fso.DeleteFile path, True
    fso.MoveFile tmp, path
    On Error GoTo 0
    Exit Sub

Fallback:
    On Error Resume Next
    If Not ts Is Nothing Then If ts.State = 1 Then ts.Close
    If Not bs Is Nothing Then If bs.State = 1 Then bs.Close
    If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    MsgBox "파일 저장 중 오류가 발생했습니다: " & Err.Description & vbCrLf & _
           "경로: " & path, vbExclamation
End Sub

Private Sub DeleteFileIfExists(ByVal path As String)
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile path, True
End Sub

Public Function GetCategoryJsonPath(ByVal cat As String) As String
    Dim nm As String
    nm = Trim$(cat)
    If nm = "" Then nm = "tasks"
    If LCase$(Right$(nm, 5)) <> ".json" Then nm = nm & ".json"
    GetCategoryJsonPath = TaskDataRootFolder() & "\" & nm
End Function


' 기존 JSONToTasks를 전부 이 내용으로 교체
Private Function JSONToTasks(ByVal json As String) As Collection
    Dim out As New Collection
    Dim raw As Variant
    Dim arr As Collection
    Dim obj As Object
    Dim t As clsTaskItem
    Dim sFrom As String, sTo As String
    Dim dF As Date, dT As Date

    On Error GoTo FAIL

    ' 후행 콤마 보정(필요 시)
    json = FixTrailingCommas(json)

    ' ParseJson: 인수 1개만!
    Set raw = JsonConverter.ParseJson(json)

    ' 루트 유형 분기 (배열 또는 { "tasks": [...] }도 허용)
    If TypeName(raw) = "Collection" Then
        Set arr = raw
    ElseIf TypeName(raw) = "Dictionary" Then
        If raw.Exists("tasks") Then
            Set arr = raw("tasks")
        Else
            Set JSONToTasks = out
            Exit Function
        End If
    Else
        Set JSONToTasks = out
        Exit Function
    End If

    ' 각 원소: {"name":..., "from":"yyyy-mm-dd", "to":"yyyy-mm-dd|null"}
    For Each obj In arr
        If TypeName(obj) = "Dictionary" Then
            sFrom = NzCStr(GetJsonStr(obj, "from"))
            sTo = NzCStr(GetJsonStr(obj, "to"))

            If (TryParseYMD(sFrom, dF) Or TryParseDate(sFrom, dF)) Then
                Set t = New clsTaskItem
                t.TaskName = NzCStr(GetJsonStr(obj, "name"))
                t.FromDate = dF
                If Len(sTo) > 0 And (TryParseYMD(sTo, dT) Or TryParseDate(sTo, dT)) Then
                    t.HasTo = True
                    t.ToDate = dT
                End If
                out.add t
            End If
        End If
    Next

    Set JSONToTasks = out
    Exit Function

FAIL:
    ' JSON 오류 시 빈 목록 반환
    Set JSONToTasks = out
End Function

'--- 유틸: 컬렉션 보조 ---
Public Sub AppendTasks(ByRef dst As Collection, ByVal src As Collection)
    If src Is Nothing Then Exit Sub
    Dim i As Long: For i = 1 To src.Count: dst.add src(i): Next
End Sub

Public Function IntersectsRange(ByVal s1 As Date, ByVal e1 As Date, _
                                ByVal s2 As Date, ByVal e2 As Date) As Boolean
    Dim t As Date
    If e1 < s1 Then t = s1: s1 = e1: e1 = t
    If e2 < s2 Then t = s2: s2 = e2: e2 = t
    IntersectsRange = Not (e1 < s2 Or e2 < s1)
End Function

'--- "보이는 구간"용: 단일 파일에서 읽고, 같은 이름의 연속/겹침을 병합 ---
Public Function LoadTasksForDateRange_File(ByVal dStart As Date, ByVal dEnd As Date) As Collection
    Dim all As Collection: Set all = LoadAllTasks_File()
    Dim filtered As New Collection
    Dim i As Long

    For i = 1 To all.Count
        Dim t As clsTaskItem, s As Date, e As Date
        Set t = all(i)
        s = t.FromDate: e = TaskEndDate(t)
        If IntersectsRange(s, e, dStart, dEnd) Then
            ' 원본을 건드리지 않도록 클론 추가
            filtered.add CloneTask(t)
        End If
    Next

    If MERGE_ADJACENT_ENABLED Then
        Set LoadTasksForDateRange_File = MergeTasksByNameAndAdjacency(filtered)
    Else
        ' 병합 없이 그대로 반환
        Set LoadTasksForDateRange_File = filtered
    End If
End Function

Public Function MergeTasksByNameAndAdjacency(ByVal tasks As Collection) As Collection
   If tasks Is Nothing Or tasks.Count = 0 Then
        Dim z As New Collection
        Set MergeTasksByNameAndAdjacency = z
        Exit Function
    End If

    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    Dim i As Long, nm As String
    For i = 1 To tasks.Count
        nm = Trim$(tasks(i).TaskName)
        If Not groups.Exists(nm) Then
            Dim c As New Collection: groups.add nm, c
        End If
        groups(nm).add tasks(i)
    Next

    Dim out As New Collection, k As Variant, merged As Collection
    For Each k In groups.keys
        Set merged = MergeOneNameGroup(groups(k), CStr(k))
        AppendTasksCloned out, merged     ' ★ 참조가 아닌 복제 추가
    Next
    Set MergeTasksByNameAndAdjacency = out
End Function

Private Function MergeOneNameGroup(ByVal grp As Collection, ByVal nameKey As String) As Collection
    Dim n As Long: n = grp.Count
    Dim arr() As clsTaskItem, i As Long, j As Long
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = grp(i): Next

    ' FromDate 오름차순
    For i = 2 To n
        Dim cur As clsTaskItem: Set cur = arr(i): j = i - 1
        Do While j >= 1 And arr(j).FromDate > cur.FromDate
            Set arr(j + 1) = arr(j): j = j - 1
        Loop
        Set arr(j + 1) = cur
    Next

    Dim out As New Collection
    If n = 0 Then Set MergeOneNameGroup = out: Exit Function

    Dim runS As Date, runE As Date
    runS = arr(1).FromDate
    runE = IIf(arr(1).HasTo, arr(1).ToDate, arr(1).FromDate)

    For i = 2 To n
        Dim s As Date, e As Date
        s = arr(i).FromDate
        e = IIf(arr(i).HasTo, arr(i).ToDate, arr(i).FromDate)

        If s <= (runE + 1) Then
            If e > runE Then runE = e
        Else
            Dim t As clsTaskItem: Set t = New clsTaskItem
            t.TaskName = nameKey: t.FromDate = runS
            t.HasTo = (runE > runS): If t.HasTo Then t.ToDate = runE
            out.add t
            runS = s: runE = e
        End If
    Next

    Dim tLast As clsTaskItem: Set tLast = New clsTaskItem
    tLast.TaskName = nameKey: tLast.FromDate = runS
    tLast.HasTo = (runE > runS): If tLast.HasTo Then tLast.ToDate = runE
    out.add tLast

    Set MergeOneNameGroup = out
End Function

'--- "현재 연도 모두 삭제" (단일 파일에서 해당 연도와 겹치는 Task만 제거 후 저장) ---
Public Sub RemoveTasksForYear_FromAll(ByVal y As Long)
    Dim all As Collection: Set all = LoadAllTasks_File()
    Dim remain As New Collection, i As Long
    Dim yS As Date, yE As Date
    yS = DateSerial(y, 1, 1): yE = DateSerial(y, 12, 31)

    For i = 1 To all.Count
        Dim t As clsTaskItem, ts As Date, te As Date
        Set t = all(i)
        ts = t.FromDate: te = TaskEndDate(t)
        If Not IntersectsRange(ts, te, yS, yE) Then
            remain.add t
        End If
    Next
    SaveAllTasks_File remain
End Sub

'==== /modTaskStoreJSON ============================================
Private Function TaskEndDate(ByVal t As clsTaskItem) As Date
    If t.HasTo Then TaskEndDate = t.ToDate Else TaskEndDate = t.FromDate
End Function

Public Function CloneTask(ByVal src As clsTaskItem) As clsTaskItem
    Dim t As New clsTaskItem
    t.TaskName = src.TaskName
    t.FromDate = src.FromDate
    t.HasTo = src.HasTo
    If src.HasTo Then t.ToDate = src.ToDate
    Set CloneTask = t
End Function

Public Sub AppendTasksCloned(ByRef dst As Collection, ByVal src As Collection)
    Dim i As Long
    If src Is Nothing Then Exit Sub
    For i = 1 To src.Count
        dst.add CloneTask(src(i))
    Next
End Sub

' === 카테고리 파일/로그 경로 ===
Public Function TaskLogFolder() As String
    Dim p As String: p = TaskDataRootFolder() & "\Log"
    EnsureFolder p
    TaskLogFolder = p
End Function

Public Function SafeFileBaseName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    s = Trim$(s)
    If Len(s) = 0 Then s = "tasks"
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next
    SafeFileBaseName = s
End Function

Public Function CategoryToFileName(ByVal cat As String) As String
    CategoryToFileName = SafeFileBaseName(IIf(Len(Trim$(cat)) = 0, "tasks", Trim$(cat))) & ".json"
End Function

Public Function CategoryJsonPath(ByVal cat As String) As String
    CategoryJsonPath = TaskDataRootFolder() & "\" & CategoryToFileName(cat)
End Function

' ===== 카테고리별 읽기/쓰기 =====
' 파일명으로 쓸 수 있도록 카테고리명 정리
Private Function SafeCategoryName(ByVal cat As String) As String
    Dim s As String: s = Trim$(cat)
    If Len(s) = 0 Then s = "tasks"   ' 기본 카테고리
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next
    SafeCategoryName = s
End Function

' 카테고리별 JSON 경로 (예: %APPDATA%\PeriodPicker\Tasks\Project1.json)
Public Function AllTasksJsonPath_Cat(ByVal cat As String) As String
    AllTasksJsonPath_Cat = TaskDataRootFolder() & "\" & SafeCategoryName(cat) & ".json"
End Function

Public Function LoadTasksForDateRange_File_Cat( _
    ByVal cat As String, ByVal dStart As Date, ByVal dEnd As Date) As Collection

    Dim all As Collection: Set all = LoadAllTasks_File_Cat(cat)
    Dim filtered As New Collection
    Dim i As Long, t As clsTaskItem, s As Date, e As Date
    For i = 1 To all.Count
        Set t = all(i)
        s = t.FromDate: e = TaskEndDate(t)
        If IntersectsRange(s, e, dStart, dEnd) Then filtered.add t
    Next
    Set LoadTasksForDateRange_File_Cat = MergeTasksByNameAndAdjacency(filtered)
End Function

Public Sub RemoveTasksForYear_FromAll_Cat(ByVal cat As String, ByVal y As Long)
    Dim all As Collection: Set all = LoadAllTasks_File_Cat(cat)
    Dim remain As New Collection
    Dim i As Long, t As clsTaskItem, s As Date, e As Date
    Dim yS As Date: yS = DateSerial(y, 1, 1)
    Dim yE As Date: yE = DateSerial(y, 12, 31)

    For i = 1 To all.Count
        Set t = all(i)
        s = t.FromDate: e = TaskEndDate(t)
        If Not IntersectsRange(s, e, yS, yE) Then remain.add t
    Next

    SaveAllTasks_File_Cat cat, remain
End Sub

' ===== JSON (VBA-JSON 기반) 대체 구현 =====
' 참고: JsonConverter.bas가 프로젝트에 포함되어 있어야 합니다.
' 필요시 [도구]-[참조]에서 "Microsoft Scripting Runtime" 체크 권장

' Array/Collection → JSON 텍스트
Private Function TasksToJsonText(ByVal tasks As Collection) As String
    Dim arr As New Collection
    Dim i As Long
    For i = 1 To IIf(tasks Is Nothing, 0, tasks.Count)
        Dim o As Object
        Set o = CreateObject("Scripting.Dictionary")
        If Len(tasks(i).TaskName) > 0 Then
            o("name") = tasks(i).TaskName
        Else
            o("name") = Null  ' null로 저장
        End If
        o("from") = Format$(tasks(i).FromDate, "yyyy-mm-dd")
        If tasks(i).HasTo Then
            o("to") = Format$(tasks(i).ToDate, "yyyy-mm-dd")
        Else
            o("to") = Null
        End If
        arr.add o
    Next
    
    TasksToJsonText = JsonConverter.ConvertToJson(arr, 2) ' ← 이름 있는 인수 사용 금지!

End Function

' --- UTF-8 BOM 제거 ---
Private Function StripUtf8Bom(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            StripUtf8Bom = Mid$(s, 2)
            Exit Function
        End If
    End If
    StripUtf8Bom = s
End Function

' --- 사전에서 대/소문자 무시 키 조회 ---
Private Function GetDictValueCI(ByVal d As Object, ByVal key As String) As Variant
    On Error GoTo EH
    If TypeName(d) <> "Dictionary" Then Exit Function
    If d.Exists(key) Then
        GetDictValueCI = d(key)
        Exit Function
    End If
    Dim k As Variant
    For Each k In d.keys
        If StrComp(CStr(k), key, vbTextCompare) = 0 Then
            GetDictValueCI = d(k)
            Exit Function
        End If
    Next
EH:
End Function

' --- Variant(사전) → clsTaskItem ---
Private Function TaskFromJsonUnknown(ByVal v As Variant) As clsTaskItem
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Dictionary" Then Exit Function

    Dim t As New clsTaskItem
    Dim nm As String, sf As String, st As String

    On Error Resume Next
    nm = NzCStr(GetDictValueCI(v, "name"))
    sf = NzCStr(GetDictValueCI(v, "from"))
    st = NzCStr(GetDictValueCI(v, "to"))
    On Error GoTo 0

    Dim dF As Date, dT As Date
    If Not (TryParseYMD(sf, dF) Or TryParseDate(sf, dF)) Then Exit Function

    t.TaskName = nm
    t.FromDate = dF
    If Len(st) > 0 Then
        If TryParseYMD(st, dT) Or TryParseDate(st, dT) Then
            t.HasTo = True
            t.ToDate = dT
        End If
    End If
    Set TaskFromJsonUnknown = t
End Function

' --- 공개 API: 카테고리 파일 로드/세이브/삭제 ---
Public Function LoadAllTasks_File_Cat(ByVal cat As String) As Collection
    Dim p As String: p = AllTasksJsonPath_Cat(cat)
    Dim json As String: json = ReadAllTextUTF8(p)
    Set LoadAllTasks_File_Cat = JsonTextToTasks(json)
End Function

Public Sub SaveAllTasks_File_Cat(ByVal cat As String, ByVal tasks As Collection)
    Dim p As String: p = AllTasksJsonPath_Cat(cat)
    Dim json As String: json = TasksToJsonText(tasks)
    SafeWriteAllTextUTF8 p, json
    ' 변경 이력 저장(로그)
    SaveTaskJsonLog cat, json
End Sub

Public Sub RemoveAllTasks_File_Cat(ByVal cat As String)
    DeleteFileIfExists AllTasksJsonPath_Cat(cat)
End Sub

' 기본 파일(API 호환 위해 유지)
Public Function LoadAllTasks_File() As Collection
    Set LoadAllTasks_File = LoadAllTasks_File_Cat("tasks")
End Function

Public Sub SaveAllTasks_File(ByVal tasks As Collection)
    SaveAllTasks_File_Cat "tasks", tasks
End Sub

Public Sub RemoveAllTasks_File()
    RemoveAllTasks_File_Cat "tasks"
End Sub

' 변경 이력 파일 저장:  %APPDATA%\PeriodPicker\Tasks\Log\{cat}_yyyymmdd_hhmmss.json
Private Sub SaveTaskJsonLog(ByVal cat As String, ByVal json As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFolder As String
    logFolder = TaskDataRootFolder() & "\Log"
    If Not fso.FolderExists(logFolder) Then fso.CreateFolder logFolder
    Dim stamp As String
    stamp = Format$(Now, "yyyymmdd_hhnnss")
    Dim fn As String
    fn = logFolder & "\" & SanitizeFileName(Nz(cat, "tasks")) & "_" & stamp & ".json"
    SafeWriteAllTextUTF8 fn, json
End Sub


Sub Test_Parse_CurrentCategory()
    Dim p$, s$, col As Collection
    'p = AllTasksJsonPath_Cat(CurrentCategoryName()) ' 프로젝트에 맞게
    p = AllTasksJsonPath_Cat("tasks") ' 프로젝트에 맞게
    s = ReadAllTextUTF8(p)
    Debug.Print "JSON length:", Len(s)
    Set col = JsonTextToTasks(s)
    Debug.Print "Parsed tasks:", col.Count
End Sub

