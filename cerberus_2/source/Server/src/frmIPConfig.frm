VERSION 5.00
Begin VB.Form frmIPConfig 
   Caption         =   "Server IP Configuration"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConfigIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtConfigPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdConfigIPOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfigIPCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Game IP"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Game Port"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmIPConfig"
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
    txtConfigIP.Text = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "IP"))
    txtConfigPort.Text = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Port"))
End Sub

Private Sub cmdConfigIPCancel_Click()
    frmLoad.Visible = False
    Unload Me
    
    End
End Sub

Private Sub cmdConfigIPOk_Click()
    PutVar App.Path & "\Data\Data.ini", "CONFIG", "IP", txtConfigIP.Text
    PutVar App.Path & "\Data\Data.ini", "CONFIG", "Port", txtConfigPort.Text
    
    Unload Me
End Sub

