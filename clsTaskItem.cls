VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== Class Module: clsTaskItem ===
Option Explicit

Private mName As String
Private mFrom As Date
Private mTo As Date
Private mHasTo As Boolean

Public Property Get TaskName() As String
    TaskName = mName
End Property

Public Property Let TaskName(ByVal v As String)
    mName = v
End Property

Public Property Get FromDate() As Date
    FromDate = mFrom
End Property

Public Property Let FromDate(ByVal v As Date)
    mFrom = v
End Property

Public Property Get HasTo() As Boolean
    HasTo = mHasTo
End Property

Public Property Let HasTo(ByVal v As Boolean)
    mHasTo = v
End Property

Public Property Get ToDate() As Date
    ToDate = mTo
End Property

Public Property Let ToDate(ByVal v As Date)
    mTo = v
End Property

