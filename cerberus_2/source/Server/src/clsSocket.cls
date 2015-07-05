VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

    ' Local variable(s) To hold Property value(s)
    Private mvarSocket As JBSOCKETSERVERLib.ISocket ' The Socket Object

    ' Custom stuff For handling the socket
    Public Sub CloseSocket()
        mvarSocket.Close
    End Sub

    Public Sub RequestRead()
        mvarSocket.RequestRead
    End Sub

    Public Sub ShutDown(how As ShutdownMethod)
        If mvarSocket Is Nothing Then Exit Sub
        Call mvarSocket.ShutDown(how)
    End Sub

    Public Sub WriteBytes(dbytes() As Byte, Optional thenShutdown As Boolean)
        Call mvarSocket.Write(dbytes, thenShutdown)
    End Sub

    Public Sub WriteString(Data As String, Optional sendAsUNICODE As Boolean, Optional thenShutdown As Boolean)
        Call mvarSocket.WriteString(Data, sendAsUNICODE, thenShutdown)
    End Sub

    Public Property Get RemoteAddress() As String
        RemoteAddress = mvarSocket.RemoteAddress.Address
    End Property

    Public Property Get RemotePort() As Long
        RemotePort = mvarSocket.RemoteAddress.Port
    End Property

    Public Property Let UserData(ByVal vData As Variant)
        mvarSocket.UserData = vData
    End Property

    Public Property Get UserData() As Variant
        UserData = mvarSocket.UserData
    End Property

    Private Sub Class_Terminate()
        Set mvarSocket = Nothing
    End Sub

    Public Property Set Socket(ByVal vData As JBSOCKETSERVERLib.ISocket)
        Set mvarSocket = vData
    End Property

    Public Property Get Socket() As JBSOCKETSERVERLib.ISocket
        Set Socket = mvarSocket
    End Property
