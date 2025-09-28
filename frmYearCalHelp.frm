VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYearCalHelp 
   Caption         =   "도움말"
   ClientHeight    =   11250
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   11143
   OleObjectBlob   =   "frmYearCalHelp.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmYearCalHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ⓒ 2025 Lim Kiljae (gninraw7@naver.com). All rights reserved. Unauthorized copying, modification, or distribution is prohibited.
Option Explicit

Private Sub UserForm_Initialize()
    Dim curLang As String
    curLang = GetLangFromReg()                 ' "K" or "E"

    ' 폼 제목/토글 초기 상태
    If curLang = "E" Then
        Me.Caption = "Year Calendar ? Help"
        tglLang.Value = True                   ' True = EN
        tglLang.Caption = "EN"
        txtHelp.text = BuildYearCalendarManual_EN()
        cmdCopy.Caption = "Copy"
        cmdClose.Caption = "Close"
    Else
        Me.Caption = "Year Calendar 도움말"
        tglLang.Value = False                  ' False = KO
        tglLang.Caption = "KO"
        txtHelp.text = BuildYearCalendarManual_KO()
        cmdCopy.Caption = "복사"
        cmdClose.Caption = "닫기"
    End If
End Sub

Private Sub tglLang_Click()
    ' 토글: True=EN, False=KO  (도움말 표시만 전환; 레지스트리는 건드리지 않음)
    If tglLang.Value Then
        tglLang.Caption = "EN"
        Me.Caption = "Year Calendar ? Help"
        txtHelp.text = BuildYearCalendarManual_EN()
        cmdCopy.Caption = "Copy"
        cmdClose.Caption = "Close"
    Else
        tglLang.Caption = "KO"
        Me.Caption = "Year Calendar 도움말"
        txtHelp.text = BuildYearCalendarManual_KO()
        cmdCopy.Caption = "복사"
        cmdClose.Caption = "닫기"
    End If
End Sub

Private Sub cmdCopy_Click()
    On Error Resume Next
    Dim dobj As New MSForms.DataObject
    dobj.SetText txtHelp.text
    dobj.PutInClipboard
    MsgBox IIf(tglLang.Value, "Copied to clipboard.", "도움말을 클립보드에 복사했습니다."), vbInformation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' 현재 언어 레지스트리 조회 (frmYearCalendar와 동일 키)
Private Function GetLangFromReg() As String
    Dim appName As String, secI18N As String
    appName = "PeriodPicker": secI18N = "I18N"
    GetLangFromReg = UCase$(GetSetting(appName, secI18N, "Lang", "K"))  ' "K" or "E"
End Function

