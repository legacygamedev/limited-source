VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Byte

Public Property Get Value() As Byte
    Value = mValue
End Property

Public Property Let Value(ByVal NewValue As Byte)
    mValue = NewValue
End Property

Public Function RTF() As String
    RTF = "\'" & HEXX(mValue)
End Function

Public Function char() As String
    char = Chr$(mValue)
End Function

Private Function HEXX(hexcode As Byte) As String
    Dim X As String
    X = "0" & Hex$(hexcode)
    HEXX = Right$(X, 2)
End Function

