VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
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

    Private WithEvents m_Server As JBSOCKETSERVERLib.Server ' The Server
Attribute m_Server.VB_VarHelpID = -1
    Public Sockets              As colSockets               ' The Socket Collection

    Private Sub Class_Initialize()
        Set m_Server = JBSOCKETSERVERLib.CreateSocketServer(GAME_PORT, GAME_IP)
        Set Sockets = New colSockets
    End Sub

    Private Sub Class_Terminate()
        Set Sockets = Nothing
        Set m_Server = Nothing
    End Sub

    Private Sub m_Server_OnConnectionClosed(ByVal Socket As JBSOCKETSERVERLib.ISocket)
        Call CloseSocket(CLng(Socket.UserData))
    End Sub

    Private Sub m_Server_OnConnectionEstablished(ByVal Socket As JBSOCKETSERVERLib.ISocket)
        Call AcceptConnection(Socket)
    End Sub

    Private Sub m_Server_OnDataReceived(ByVal Socket As JBSOCKETSERVERLib.ISocket, ByVal Data As JBSOCKETSERVERLib.IData)
        Call IncomingData(Socket, Data)
    End Sub

    Public Sub StartListening()
        m_Server.StartListening
    End Sub

    Public Sub StopListening()
        m_Server.StopListening
    End Sub

    Public Property Get LocalAddress() As String
        LocalAddress = m_Server.LocalAddress.Address
    End Property

    Public Property Get LocalPort() As Long
        LocalPort = m_Server.LocalAddress.Port
    End Property
