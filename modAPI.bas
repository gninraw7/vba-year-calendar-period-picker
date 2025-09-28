Attribute VB_Name = "modAPI"
Option Explicit

' ====== 공휴일(휴무일) API 기본 ======
Private Const BASE_URL As String = _
  "https://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo"
Private Const SVC_KEY As String = "02d66e72d83dde0583541dbf6d5af6fa15ed9c8214712f46c5c0696b89ec93cf"

' item 필드: locdate(yyyymmdd), dateName(이름), isHoliday(Y/N)
Private Function FetchHolidaysOneYear(ByVal y As Long) As Collection
    Dim url As String
    url = BASE_URL & "?ServiceKey=" & SVC_KEY & "&solYear=" & CStr(y) & "&numOfRows=100"

    Dim xhr As Object: Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", url, False
    xhr.send

    Dim col As New Collection
    If xhr.Status <> 200 Then
        Set FetchHolidaysOneYear = col
        Exit Function
    End If

    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
    xml.async = False
    xml.LoadXML xhr.responseText

    Dim items As Object, it As Object
    Set items = xml.getElementsByTagName("item")
    Dim d As String, nm As String

    For Each it In items
        d = it.SelectSingleNode("locdate").text  ' yyyymmdd
        nm = it.SelectSingleNode("dateName").text
        ' yyyy-mm-dd로 정규화
        d = Left$(d, 4) & "-" & Mid$(d, 5, 2) & "-" & Right$(d, 2)

        Dim rec(1) As String
        rec(0) = d: rec(1) = nm
        col.add rec
    Next

    Set FetchHolidaysOneYear = col
End Function

' ====== 지정한 연도의 공휴일 정보를 API로 불러오기 → 새 시트에 출력 ======
Public Sub ExportKRPublicHolidays(y As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    On Error Resume Next
    ws.name = "API_" & Format(y, "####") & "_공휴일_" & Format(Now, "yymmdd_hhnnss")
    On Error GoTo 0

    ws.Range("A1").Value = "휴일"
    ws.Range("B1").Value = "휴일명"

    Dim r As Long
    r = 2
    Dim bag As New Collection, one As Collection, i As Long

    ' 1) 지정한 연도의 공휴일 불러오기
    Set one = FetchHolidaysOneYear(y)
    For i = 1 To one.Count
        ws.Cells(r, 1).Value = one(i)(0) ' yyyy-mm-dd
        ws.Cells(r, 2).Value = one(i)(1) ' 이름
        r = r + 1
    Next

    ' 2) 정렬(날짜 오름차순)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    With ws.Sort
        .SortFields.Clear
        .SortFields.add ws.Range("A2:A" & lastRow), xlSortOnValues, xlAscending
        .SetRange ws.Range("A1:B" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' 3) 표시/서식
    ws.Columns(1).NumberFormat = "yyyy-mm-dd"
    ws.Columns("A:B").AutoFit

    MsgBox "API 공휴일 정보: " & (lastRow - 1) & "건을 '" & ws.name & "' 시트에 출력했습니다.", vbInformation
End Sub




