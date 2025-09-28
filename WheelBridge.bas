Attribute VB_Name = "WheelBridge"
'Option Explicit
'
'#If VBA7 Then
'    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'#Else
'    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'#End If
'
'' 이 모듈은 "휠 신호 → 대상 폼" 전달만 담당
'Private mSink As Object   ' frmYearCalendar 같은 대상 폼을 보관
'
'Public Sub RegisterWheelSink(ByVal sink As Object)
'    Set mSink = sink
'End Sub
'
'Public Sub UnregisterWheelSink(ByVal sink As Object)
'    On Error Resume Next
'    If Not mSink Is Nothing Then
'        If mSink Is sink Then Set mSink = Nothing
'    End If
'End Sub
'
'' MouseOverControl 쪽에서 호출: RequestDy(위/아래 방향)만 사용
'Public Sub FireWheel(ByVal RequestDy As Single)
'    If mSink Is Nothing Then Exit Sub
'
'    ' 방향(+/-)만 필요
'    Dim dir As Long: dir = Sgn(RequestDy)
'
'    ' 보정: Shift/Ctrl 상태(연도 가산폭 결정에 사용)
'    Const VK_SHIFT As Long = &H10
'    Const VK_CONTROL As Long = &H11
'    Dim isShift As Boolean, isCtrl As Boolean
'    isShift = (GetKeyState(VK_SHIFT) And &H8000) <> 0
'    isCtrl = (GetKeyState(VK_CONTROL) And &H8000) <> 0
'
'    On Error Resume Next
'    ' 대상 폼은 Public Sub OnMouseWheel(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)를 구현
'    CallByName mSink, "OnMouseWheel", VbMethod, dir, isShift, isCtrl
'End Sub
'
'' ==== 추가 ====
'Public Sub FireWheelFromHook(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)
'    If mSink Is Nothing Then Exit Sub
'    On Error Resume Next
'    CallByName mSink, "OnMouseWheel", VbMethod, dir, isShift, isCtrl
'End Sub
'
