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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSockets"
Attribute VB_Ext_KEY = "Member0" ,"clsSockets"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable to hold collection
Private mCol As Collection
Public Function Add(Key As String, Optional sKey As String) As clsSocket
Attribute Add.VB_Description = "Create a new object in the collection."
  'create a new object
  Dim objNewMember As clsSocket
  Set objNewMember = New clsSocket

  'set the properties passed into the method
  objNewMember.Key = Key
  If Len(sKey) = 0 Then
    mCol.Add objNewMember
  Else
    mCol.Add objNewMember, sKey
  End If
  
  'return the object created
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function
Public Property Get Item(vntIndexKey As Variant) As clsSocket
Attribute Item.VB_Description = "Used when referencing an element in the collection. vntIndexKey contains either the Index or Key to the collection. Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)"
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
Attribute Count.VB_Description = "Used when retrieving the number of elements in the collection. Syntax: Debug.Print x.Count"
  Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "Used when removing an element from the collection. vntIndexKey contains either the Index or Key. Syntax: x.Remove(xyz)"
  Call mCol(vntIndexKey).Shutdown(ShutdownBoth)
  mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "This property allows you to enumerate this collection with the For...Each syntax."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
  'destroys collection when this class is terminated
  Set mCol = Nothing
End Sub
