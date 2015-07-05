VERSION 5.00
Begin VB.Form frmIPconfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfigIPCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfigIPOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtConfigPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtConfigIP 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox chkConfigAlways 
      Caption         =   "Always use this IP address"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Game Port"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Game IP"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmIPconfig"
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

Private Sub Form_Load()
    txtConfigIP.Text = Trim(GetVar(App.Path & "\config.ini", "IPCONFIG", "IP"))
    txtConfigPort.Text = Trim(GetVar(App.Path & "\config.ini", "IPCONFIG", "Port"))
    chkConfigAlways.Value = GetVar(App.Path & "\config.ini", "IPCONFIG", "Always")
End Sub

Private Sub chkConfigAlways_Click()
    If chkConfigAlways.Value = Checked Then
        PutVar App.Path & "\config.ini", "IPCONFIG", "Always", 1
    Else
        PutVar App.Path & "\config.ini", "IPCONFIG", "Always", 0
    End If
End Sub

Private Sub cmdConfigIPCancel_Click()
    frmSendGetData.Visible = False
    Unload Me
    
    End
End Sub

Private Sub cmdConfigIPOk_Click()
    PutVar App.Path & "\config.ini", "IPCONFIG", "IP", txtConfigIP.Text
    PutVar App.Path & "\config.ini", "IPCONFIG", "Port", txtConfigPort.Text
    GAME_IP = txtConfigIP.Text
    GAME_PORT = Val(txtConfigPort.Text)
    
    Unload Me
End Sub

