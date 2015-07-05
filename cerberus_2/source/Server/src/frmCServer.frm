VERSION 5.00
Begin VB.Form frmCServer 
   Caption         =   "Cerberus Server"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGameAI 
      Interval        =   1000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtText 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   10215
   End
   Begin VB.TextBox txtChat 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   10215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call GlobalMsg(txtChat.Text, Black)
        Call TextAdd(frmCServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
End Sub


' ************
' ** Timers **
' ************

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    If Secs = 30 Or Secs = 20 Or Secs = 10 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    End If
    Call TextAdd(frmCServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    Secs = Secs - 2
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

' ******************
' ** Menu Options **
' ******************

Private Sub mnuServerLog_Click()
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
End Sub

Private Sub mnuShutdown_Click()
    tmrShutdown.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Call DestroyServer
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(frmCServer.txtText, "All classes reloaded.", True)
End Sub

