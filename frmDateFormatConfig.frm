VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateFormatConfig 
   Caption         =   "설정"
   ClientHeight    =   4763
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   7227
   OleObjectBlob   =   "frmDateFormatConfig.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmDateFormatConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit


Private Sub UserForm_Initialize()
    Dim lang As String
    lang = GetSetting("PeriodPicker", "I18N", "Lang", "K")

    ' 날짜 형식(공용)
    Dim defDateS As String
    defDateS = GetSetting("PeriodPicker", "Format", "Date", _
                IIf(UCase$(lang) = "E", "yyyy-mm-dd", "yyyy-mm-dd"))
    txtDateFmt.text = defDateS

    ' --- 월 타이틀(언어별) ---
    Dim tk As String, te As String, legacy As String
    tk = GetSetting("PeriodPicker", "Format", "MonthTitle_K", "yyyy""년"" m""월""")
    te = GetSetting("PeriodPicker", "Format", "MonthTitle_E", "mmmm yyyy")
'    legacy = GetSetting("PeriodPicker", "Format", "MonthTitle", "")
'    If legacy <> "" Then
'        If tk = "yyyy""년"" m""월""" Then tk = legacy
'        If te = "mmmm yyyy" Then te = legacy
'    End If
    txtTitleFmtK.text = tk
    txtTitleFmtE.text = te
    PreviewAll
    
    Dim ws As String, wn As String
    ws = GetSetting("PeriodPicker", "I18N", "WeekStart", "Sun")
    wn = GetSetting("PeriodPicker", "I18N", "WeekNameStyle", "Short")
    chkMonFirst.Value = (UCase$(ws) = "MON")
    optWkFull.Value = (UCase$(wn) = "FULL")
    optWkShort.Value = Not optWkFull.Value
    
End Sub

Private Sub txtDateFmt_Change()
    PreviewAll
End Sub

Private Sub txtTitleFmtK_Change(): PreviewAll: End Sub
Private Sub txtTitleFmtE_Change(): PreviewAll: End Sub

Private Sub PreviewAll()
    On Error Resume Next
    lblEx1.Caption = "Date preview: " & Format(Date, NzEmptyToDefault(txtDateFmt.text, "yyyy-mm-dd"))
    lblExTitleK.Caption = "K title: " & Format(DateSerial(Year(Date), Month(Date), 1), _
                        NzEmptyToDefault(txtTitleFmtK.text, "yyyy""년"" m""월"""))
    lblExTitleE.Caption = "E title: " & Format(DateSerial(Year(Date), Month(Date), 1), _
                        NzEmptyToDefault(txtTitleFmtE.text, "mmmm yyyy"))
End Sub

Private Function NzEmptyToDefault(ByVal s As String, ByVal d As String) As String
    If Len(Trim$(s)) = 0 Then NzEmptyToDefault = d Else NzEmptyToDefault = s
End Function

Private Sub btnOK_Click()
    On Error GoTo bad
    Dim d As String, tk As String, te As String
    d = txtDateFmt.text: tk = txtTitleFmtK.text: te = txtTitleFmtE.text
    ' 간단 검증
    Dim t_x As String
    t_x = Format(Date, d)
    t_x = Format(DateSerial(Year(Date), Month(Date), 1), tk)
    t_x = Format(DateSerial(Year(Date), Month(Date), 1), te)

    SaveSetting "PeriodPicker", "Format", "Date", d
    SaveSetting "PeriodPicker", "Format", "MonthTitle_K", tk
    SaveSetting "PeriodPicker", "Format", "MonthTitle_E", te
'    ' 선택 사항: 구버전 키도 마지막으로 한 번 동기화
'    SaveSetting "PeriodPicker", "Format", "MonthTitle", _
'        IIf(UCase$(GetSetting("PeriodPicker", "I18N", "Lang", "K")) = "E", te, tk)

    SaveSetting "PeriodPicker", "I18N", "WeekStart", IIf(chkMonFirst.Value, "Mon", "Sun")
    SaveSetting "PeriodPicker", "I18N", "WeekNameStyle", IIf(optWkFull.Value, "Full", "Short")
    
    Unload Me
    Exit Sub
bad:
    MsgBox "형식이 올바르지 않습니다." & vbCrLf & _
           "예) yyyy-mm-dd / mmmm yyyy / yyyy""년"" m""월""", vbExclamation
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


