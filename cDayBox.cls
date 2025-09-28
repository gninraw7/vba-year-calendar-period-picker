VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDayBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents tb    As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents cb    As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1
Public WithEvents lb    As MSForms.ListBox
Attribute lb.VB_VarHelpID = -1
Public WithEvents btn   As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents opt   As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Public WithEvents spn   As MSForms.SpinButton
Attribute spn.VB_VarHelpID = -1
Public WithEvents lbl   As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Public WithEvents chk   As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1

Public Host As Object   ' frmYearCalendar 같은 호스트폼

Public Sub Hook(ByVal ctrl As Object, ByVal owner As frmYearCalendar)
    Set Host = owner
    Select Case TypeName(ctrl)
        Case "TextBox":         Set tb = ctrl
        Case "ComboBox":        Set cb = ctrl
        Case "ListBox":         Set lb = ctrl
        Case "CommandButton":   Set btn = ctrl
        Case "OptionButton":    Set opt = ctrl
        Case "SpinButton":      Set spn = ctrl
        Case "Label":           Set lbl = ctrl
        Case "CheckBox":        Set chk = ctrl
    End Select
End Sub

Private Sub tb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    On Error Resume Next
    Host.HandleDayMouse tb, Button  ' 좌=1 → From, 우=2 → To
End Sub

Private Sub tb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Host.HandleDayDbl tb            ' 더블클릭 즉시 반환
End Sub

Private Sub tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub cb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub lb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub opt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub spn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub chk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Host.RouteKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    On Error Resume Next
    ' 월 타이틀 레이블만 처리 (예: "lblMonth1"~"lblMonth12")
    If Left$(lbl.name, 8) = "lblMonth" Then
        Host.HandleMonthLabelMouse lbl, Button
    End If
End Sub

' cDayBox.cls
Private Sub lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Left$(lbl.name, 8) = "lblMonth" Then
        Host.HandleMonthLabelDouble lbl
    End If
End Sub

