VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSockets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable To hold Collection
Private mCol As Collection

Public Function Add(Optional sKey As String) As clsSocket
    ' Create New Object
    Dim objNewMember As clsSocket
    Set objNewMember = New clsSocket
   
    ' Set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   
    ' Return the Object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsSocket
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Call mCol(vntIndexKey).ShutDown(ShutdownBoth)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    ' Creates the Collection when this Is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    ' Destroys Collection when this class Is terminated
    Set mCol = Nothing
End Sub
