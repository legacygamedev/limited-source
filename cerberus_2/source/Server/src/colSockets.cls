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
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

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
