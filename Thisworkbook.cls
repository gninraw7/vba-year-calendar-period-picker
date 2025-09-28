VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thisworkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    SeedHolidays_InitialInstall
    Gen_PeriodPicker_CommandBar
End Sub
Private Sub Workbook_Open()
    Gen_PeriodPicker_CommandBar
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Delete_PeriodPicker_Menu
End Sub
Private Sub Workbook_AddinUninstall()
    Delete_PeriodPicker_Menu
End Sub


