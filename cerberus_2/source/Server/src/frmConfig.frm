VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Configuration"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfigOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdConfigDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtGameName 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Text            =   "Cerberus Default"
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox txtWebsite 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Text            =   "http://www.webaddress.com"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Text            =   "7000"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMaxPlayers 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Text            =   "10"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtMaxMaps 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Text            =   "255"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtMaxNpcs 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Text            =   "50"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtMaxMapItems 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Text            =   "20"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtMaxSpells 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   "50"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtMaxSkills 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "50"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtMaxQuests 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "50"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtMaxShops 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Text            =   "15"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtMaxGuilds 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Text            =   "20"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtMaxMembers 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Text            =   "10"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox txtMaxItems 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "50"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdConfigCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdConfigMax 
      Caption         =   "Max"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox chkIP 
      Caption         =   "Always use this IP"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Website URL"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Game Port"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Maximum Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Max 
      Caption         =   "Max Players"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Max Maps"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Max NPC's"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Max Map Items"
      Height          =   255
      Left            =   2880
      TabIndex        =   28
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Max Spells"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Max Skills"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Max Quests"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Max Shops"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Max Guilds"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Max Members"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Max Items"
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Game IP"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfig"
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

Private Sub cmdConfigCancel_Click()
    frmLoad.Visible = False
    Unload Me
    
    End
End Sub

Private Sub cmdConfigDefault_Click()
    txtMaxPlayers.Text = "10"
    txtMaxItems.Text = "50"
    txtMaxNpcs.Text = "50"
    txtMaxShops.Text = "15"
    txtMaxSpells.Text = "50"
    txtMaxSkills.Text = "50"
    txtMaxMaps.Text = "255"
    txtMaxMapItems.Text = "20"
    txtMaxGuilds.Text = "20"
    txtMaxMembers.Text = "10"
    'txtMaxEmoticons.Text = "10"
    txtMaxQuests.Text = "50"
    'txtMaxLevel.Text = "500"
End Sub

Private Sub cmdConfigMax_Click()
    txtMaxPlayers.Text = "50"
    txtMaxItems.Text = "250"
    txtMaxNpcs.Text = "250"
    txtMaxShops.Text = "50"
    txtMaxSpells.Text = "250"
    txtMaxSkills.Text = "250"
    txtMaxMaps.Text = "1000"
    txtMaxMapItems.Text = "25"
    txtMaxGuilds.Text = "30"
    txtMaxMembers.Text = "20"
    'txtMaxEmoticons.Text = "10"
    txtMaxQuests.Text = "250"
    'txtMaxLevel.Text = "500"
End Sub

Private Sub cmdConfigOk_Click()
        PutVar App.Path & "\Data\Data.ini", "CONFIG", "GameName", txtGameName.Text
        PutVar App.Path & "\Data\Data.ini", "CONFIG", "WebSite", txtWebsite.Text
        PutVar App.Path & "\Data\Data.ini", "CONFIG", "IP", txtIP.Text
        If chkIP.Value = Checked Then
            PutVar App.Path & "\Data\Data.ini", "CONFIG", "Always", 1
        Else
            PutVar App.Path & "\Data\Data.ini", "CONFIG", "Always", 0
        End If
        PutVar App.Path & "\Data\Data.ini", "CONFIG", "Port", txtPort.Text
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "HPRegen", 1
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "MPRegen", 1
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "SPRegen", 1
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "Scrolling", 1
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "AutoTurn", 1
        'PutVar App.Path & "\Data\Data.ini", "CONFIG", "Scripting", 1
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_PLAYERS", txtMaxPlayers.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_ITEMS", txtMaxItems.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_NPCS", txtMaxNpcs.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_SHOPS", txtMaxShops.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_SPELLS", txtMaxSpells.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_SKILLS", txtMaxSkills.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_MAPS", txtMaxMaps.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_MAP_ITEMS", txtMaxMapItems.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_GUILDS", txtMaxGuilds.Text
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_GUILD_MEMBERS", txtMaxMembers.Text
        'PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_EMOTICONS", 10
        PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_QUESTS", txtMaxQuests.Text
        'PutVar App.Path & "\Data\Data.ini", "MAX", "MAX_LEVEL", 500
        
        Unload Me
End Sub

