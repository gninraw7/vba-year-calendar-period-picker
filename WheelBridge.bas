Attribute VB_Name = "WheelBridge"
'Option Explicit
'
'#If VBA7 Then
'    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'#Else
'    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'#End If
'
'' �� ����� "�� ��ȣ �� ��� ��" ���޸� ���
'Private mSink As Object   ' frmYearCalendar ���� ��� ���� ����
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
'' MouseOverControl �ʿ��� ȣ��: RequestDy(��/�Ʒ� ����)�� ���
'Public Sub FireWheel(ByVal RequestDy As Single)
'    If mSink Is Nothing Then Exit Sub
'
'    ' ����(+/-)�� �ʿ�
'    Dim dir As Long: dir = Sgn(RequestDy)
'
'    ' ����: Shift/Ctrl ����(���� ������ ������ ���)
'    Const VK_SHIFT As Long = &H10
'    Const VK_CONTROL As Long = &H11
'    Dim isShift As Boolean, isCtrl As Boolean
'    isShift = (GetKeyState(VK_SHIFT) And &H8000) <> 0
'    isCtrl = (GetKeyState(VK_CONTROL) And &H8000) <> 0
'
'    On Error Resume Next
'    ' ��� ���� Public Sub OnMouseWheel(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)�� ����
'    CallByName mSink, "OnMouseWheel", VbMethod, dir, isShift, isCtrl
'End Sub
'
'' ==== �߰� ====
'Public Sub FireWheelFromHook(ByVal dir As Long, ByVal isShift As Boolean, ByVal isCtrl As Boolean)
'    If mSink Is Nothing Then Exit Sub
'    On Error Resume Next
'    CallByName mSink, "OnMouseWheel", VbMethod, dir, isShift, isCtrl
'End Sub
'
