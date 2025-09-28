Attribute VB_Name = "modTaskStoreJSON"
'==== modTaskStoreJSON (�߰�/��ü) =================================
Option Explicit

' ���� ��� ����(���� �⺻��: False)
Private Const MERGE_ADJACENT_ENABLED As Boolean = False

Private Sub InitJsonOptions()
    With JsonConverter.JsonOptions
        .UseDoubleForLargeNumbers = True
        .AllowUnquotedKeys = False
        .EscapeSolidus = False
    End With
End Sub

Private Function FixTrailingCommas(ByVal s As String) As String
    ' �ſ� �ܼ��� ����: ",]" -> "]", ",}" -> "}"
    FixTrailingCommas = Replace(Replace(s, ",]", "]"), ",}", "}")
End Function

Private Function GetJsonStr(ByVal d As Object, ByVal k As String) As String
    On Error Resume Next
    If Not d Is Nothing Then
        If d.Exists(k) Then GetJsonStr = CStr(d(k))
    End If
End Function

' UTF-8 BOM ���� +(����) ���� ��ó��
Private Function NormalizeJson(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2) ' UTF-8 BOM ����
    End If
    ' �ּ�/Ʈ���ϸ� �޸� ��� �� �ϹǷ�, ���� �ʿ��� �ùٸ� JSON�� ���鵵�� �����ϼ���.
    NormalizeJson = s
End Function

' Dictionary/Collection�� �����ϰ� clsTaskItem���� ��ȯ
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
    ' �� �׸��� Ʋ���� ��ü�� ���
End Sub

' Dictionary���� Ű ���� ����
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

' === ����: �ؽ�Ʈ �� Collection(clsTaskItem) ===
' modTaskStoreJSON
Private Function JsonTextToTasks(ByVal json As String) As Collection
    On Error GoTo EH

    ' �ʿ� �� BOM/������ ��ǥ ����(�̹� �����ø� �״�� �μ���)
    json = NormalizeJson(json)

    ' <<< �߿�: �� ������ ��(.) �Ӽ����θ� ���� ���� >>>
    With JsonConverter.JsonOptions
        .UseDoubleForLargeNumbers = True   ' �Ǵ� False, ������Ʈ ��å�� �°�
        .AllowUnquotedKeys = False
        .EscapeSolidus = False
    End With

    Dim root As Variant
    Set root = JsonConverter.ParseJson(json)  ' �μ� 1����!

    Dim out As New Collection
    Dim i As Long

    If TypeName(root) = "Collection" Then
        For i = 1 To root.Count
            Dim t As clsTaskItem
            Set t = TaskFromDict(root(i))
            If Not t Is Nothing Then out.add t
        Next
    ElseIf TypeName(root) = "Dictionary" Then
        ' Ȥ�� {"tasks":[...]} ������ ����Ϸ���
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
    MsgBox "JSON �Ľ� ���� (" & Err.Number & "): " & Err.Description, vbExclamation
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

' ���� ����: %APPDATA%\PeriodPicker\Tasks\tasks.json
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

'--- ���� I/O (UTF-8, ���� ����) ---
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

    ' 1) Text ��Ʈ���� UTF-8�� ���
    With ts
        .Type = 2              ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText text, 0     ' adWriteLine=1 �ƴ�
        .Position = 0
    End With

    ' 2) Binary ��Ʈ������ ���� �� ����
    With bs
        .Type = 1              ' adTypeBinary
        .Open
        ts.CopyTo bs
        ts.Close               ' �ҽ� ���� �ݾ� ��� ����
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
    MsgBox "���� ���� �� ������ �߻��߽��ϴ�: " & Err.Description & vbCrLf & _
           "���: " & path, vbExclamation
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


' ���� JSONToTasks�� ���� �� �������� ��ü
Private Function JSONToTasks(ByVal json As String) As Collection
    Dim out As New Collection
    Dim raw As Variant
    Dim arr As Collection
    Dim obj As Object
    Dim t As clsTaskItem
    Dim sFrom As String, sTo As String
    Dim dF As Date, dT As Date

    On Error GoTo FAIL

    ' ���� �޸� ����(�ʿ� ��)
    json = FixTrailingCommas(json)

    ' ParseJson: �μ� 1����!
    Set raw = JsonConverter.ParseJson(json)

    ' ��Ʈ ���� �б� (�迭 �Ǵ� { "tasks": [...] }�� ���)
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

    ' �� ����: {"name":..., "from":"yyyy-mm-dd", "to":"yyyy-mm-dd|null"}
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
    ' JSON ���� �� �� ��� ��ȯ
    Set JSONToTasks = out
End Function

'--- ��ƿ: �÷��� ���� ---
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

'--- "���̴� ����"��: ���� ���Ͽ��� �а�, ���� �̸��� ����/��ħ�� ���� ---
Public Function LoadTasksForDateRange_File(ByVal dStart As Date, ByVal dEnd As Date) As Collection
    Dim all As Collection: Set all = LoadAllTasks_File()
    Dim filtered As New Collection
    Dim i As Long

    For i = 1 To all.Count
        Dim t As clsTaskItem, s As Date, e As Date
        Set t = all(i)
        s = t.FromDate: e = TaskEndDate(t)
        If IntersectsRange(s, e, dStart, dEnd) Then
            ' ������ �ǵ帮�� �ʵ��� Ŭ�� �߰�
            filtered.add CloneTask(t)
        End If
    Next

    If MERGE_ADJACENT_ENABLED Then
        Set LoadTasksForDateRange_File = MergeTasksByNameAndAdjacency(filtered)
    Else
        ' ���� ���� �״�� ��ȯ
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
        AppendTasksCloned out, merged     ' �� ������ �ƴ� ���� �߰�
    Next
    Set MergeTasksByNameAndAdjacency = out
End Function

Private Function MergeOneNameGroup(ByVal grp As Collection, ByVal nameKey As String) As Collection
    Dim n As Long: n = grp.Count
    Dim arr() As clsTaskItem, i As Long, j As Long
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = grp(i): Next

    ' FromDate ��������
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

'--- "���� ���� ��� ����" (���� ���Ͽ��� �ش� ������ ��ġ�� Task�� ���� �� ����) ---
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

' === ī�װ� ����/�α� ��� ===
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

' ===== ī�װ��� �б�/���� =====
' ���ϸ����� �� �� �ֵ��� ī�װ��� ����
Private Function SafeCategoryName(ByVal cat As String) As String
    Dim s As String: s = Trim$(cat)
    If Len(s) = 0 Then s = "tasks"   ' �⺻ ī�װ�
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next
    SafeCategoryName = s
End Function

' ī�װ��� JSON ��� (��: %APPDATA%\PeriodPicker\Tasks\Project1.json)
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

' ===== JSON (VBA-JSON ���) ��ü ���� =====
' ����: JsonConverter.bas�� ������Ʈ�� ���ԵǾ� �־�� �մϴ�.
' �ʿ�� [����]-[����]���� "Microsoft Scripting Runtime" üũ ����

' Array/Collection �� JSON �ؽ�Ʈ
Private Function TasksToJsonText(ByVal tasks As Collection) As String
    Dim arr As New Collection
    Dim i As Long
    For i = 1 To IIf(tasks Is Nothing, 0, tasks.Count)
        Dim o As Object
        Set o = CreateObject("Scripting.Dictionary")
        If Len(tasks(i).TaskName) > 0 Then
            o("name") = tasks(i).TaskName
        Else
            o("name") = Null  ' null�� ����
        End If
        o("from") = Format$(tasks(i).FromDate, "yyyy-mm-dd")
        If tasks(i).HasTo Then
            o("to") = Format$(tasks(i).ToDate, "yyyy-mm-dd")
        Else
            o("to") = Null
        End If
        arr.add o
    Next
    
    TasksToJsonText = JsonConverter.ConvertToJson(arr, 2) ' �� �̸� �ִ� �μ� ��� ����!

End Function

' --- UTF-8 BOM ���� ---
Private Function StripUtf8Bom(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            StripUtf8Bom = Mid$(s, 2)
            Exit Function
        End If
    End If
    StripUtf8Bom = s
End Function

' --- �������� ��/�ҹ��� ���� Ű ��ȸ ---
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

' --- Variant(����) �� clsTaskItem ---
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

' --- ���� API: ī�װ� ���� �ε�/���̺�/���� ---
Public Function LoadAllTasks_File_Cat(ByVal cat As String) As Collection
    Dim p As String: p = AllTasksJsonPath_Cat(cat)
    Dim json As String: json = ReadAllTextUTF8(p)
    Set LoadAllTasks_File_Cat = JsonTextToTasks(json)
End Function

Public Sub SaveAllTasks_File_Cat(ByVal cat As String, ByVal tasks As Collection)
    Dim p As String: p = AllTasksJsonPath_Cat(cat)
    Dim json As String: json = TasksToJsonText(tasks)
    SafeWriteAllTextUTF8 p, json
    ' ���� �̷� ����(�α�)
    SaveTaskJsonLog cat, json
End Sub

Public Sub RemoveAllTasks_File_Cat(ByVal cat As String)
    DeleteFileIfExists AllTasksJsonPath_Cat(cat)
End Sub

' �⺻ ����(API ȣȯ ���� ����)
Public Function LoadAllTasks_File() As Collection
    Set LoadAllTasks_File = LoadAllTasks_File_Cat("tasks")
End Function

Public Sub SaveAllTasks_File(ByVal tasks As Collection)
    SaveAllTasks_File_Cat "tasks", tasks
End Sub

Public Sub RemoveAllTasks_File()
    RemoveAllTasks_File_Cat "tasks"
End Sub

' ���� �̷� ���� ����:  %APPDATA%\PeriodPicker\Tasks\Log\{cat}_yyyymmdd_hhmmss.json
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
    'p = AllTasksJsonPath_Cat(CurrentCategoryName()) ' ������Ʈ�� �°�
    p = AllTasksJsonPath_Cat("tasks") ' ������Ʈ�� �°�
    s = ReadAllTextUTF8(p)
    Debug.Print "JSON length:", Len(s)
    Set col = JsonTextToTasks(s)
    Debug.Print "Parsed tasks:", col.Count
End Sub

