VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tmrName As String
Private lngInterval As Long
Private tmrLeft As Long
Public Property Let Name(ByVal CName As String)
    tmrName = CName
End Property
Public Property Let Interval(ByVal CInterval As Long)
    lngInterval = CInterval
End Property
Public Property Get Name() As String
    Name = tmrName
End Property
Public Property Get Interval() As Long
    Interval = lngInterval
End Property
Public Property Get tmrWait() As Long
    tmrWait = tmrLeft
End Property
Public Property Let tmrWait(ByVal CTime As Long)
    tmrLeft = CTime
End Property
