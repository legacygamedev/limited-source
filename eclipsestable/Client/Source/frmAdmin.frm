VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Administration Panel"
   ClientHeight    =   3390
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   3960
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "frmAdmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "Close Admin Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5953
      _Version        =   393216
      Tab             =   2
      TabHeight       =   353
      TabMaxWidth     =   1235
      BackColor       =   -2147483648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Player"
      TabPicture(0)   =   "frmAdmin.frx":0FC2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblPlayerName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblValue"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnSetAccess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnKick"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnWarpMeTo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnBan"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnSetSprite"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtValue"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPlayerName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnWarpToMe"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "World"
      TabPicture(1)   =   "frmAdmin.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnSetSnow"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnSetRain"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnSetThunder"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "btnSetNone"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "btnWarpTo"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "btnRespawn"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btnLocation"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtMap"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblWeather"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblMapNumber"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Editor"
      TabPicture(2)   =   "frmAdmin.frx":0FFA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "btnEditSpell"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "btnEditItem"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnEditShops"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "btnEditNPC"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnEditArrow"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "btnEditEmoticon"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "btnEditElement"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "btnEditMap"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "CmdEditMain"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "TxtScriptName"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.TextBox TxtScriptName 
         Height          =   285
         Left            =   2040
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdEditMain 
         Caption         =   "Edit Server Script"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpToMe 
         Caption         =   "Warp To Me"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton btnEditMap 
         Caption         =   "Edit Map"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         MaskColor       =   &H8000000D&
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton btnEditElement 
         Caption         =   "Edit Elements"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton btnEditEmoticon 
         Caption         =   "Edit Emoticons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton btnEditArrow 
         Caption         =   "Edit Arrows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnEditNPC 
         Caption         =   "Edit NPCs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnEditShops 
         Caption         =   "Edit Shops"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton btnEditItem 
         Caption         =   "Edit Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton btnEditSpell 
         Caption         =   "Edit Spells"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton btnSetSnow 
         Caption         =   "Snow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   20
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSetRain 
         Caption         =   "Rain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton btnSetThunder 
         Caption         =   "Thunder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSetNone 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpTo 
         Caption         =   "Warp To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnRespawn 
         Caption         =   "Respawn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton btnLocation 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtMap 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtPlayerName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72960
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnSetSprite 
         Caption         =   "Set Sprite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton btnBan 
         Caption         =   "Ban Player"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpMeTo 
         Caption         =   "Warp Me To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnKick 
         Caption         =   "Kick Player"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton btnSetAccess 
         Caption         =   "Set Access"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblWeather 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weather:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   29
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label lblValue 
         Caption         =   "Set Value:"
         Height          =   255
         Left            =   -72960
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblMapNumber 
         Caption         =   "Map Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPlayerName 
         Caption         =   "Player Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEditMap_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestEditMap
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnSetSprite_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            If LenB(txtValue.Text) <> 0 Then
                If IsNumeric(txtValue.Text) Then
                    If Not Val(txtValue.Text) < 1 Then
                        Call SendSetPlayerSprite(txtPlayerName.Text, txtValue.Text)
                    End If
                End If
            End If
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnBan_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call SendBan(txtPlayerName.Text)
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditItem_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditItem
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditShops_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditShop
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditSpell_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditSpell
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnKick_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call SendKick(txtPlayerName.Text)
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnLocation_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        BLoc = Not BLoc
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnRespawn_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendMapRespawn
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpMeTo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call WarpMeTo(txtPlayerName.Text)
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpTo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Len(txtMap.Text) <> 0 Then
            If GetPlayerMap(MyIndex) <> Val(txtMap.Text) Then
                Call WarpTo(Val(txtMap.Text), GetPlayerX(MyIndex), GetPlayerY(MyIndex))
            Else
                Call AddText("You are already on this map. You cannot warp to it.", BRIGHTRED)
            End If
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpToMe_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call WarpToMe(txtPlayerName.Text)
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub PlayerInfo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName) <> 0 Then
            Call SendData("getstats" & SEP_CHAR & txtPlayerName.Text & END_CHAR)
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditArrow_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditArrow
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditEmoticon_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditEmoticon
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditNPC_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditNPC
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditElement_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditElement
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnSetAccess_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
        If LenB(txtPlayerName.Text) <> 0 Then
            If LenB(txtValue.Text) <> 0 Then
                If Val(txtValue.Text) < 0 Or Val(txtValue.Text) > 5 Then
                    Call AddText("Valid access range is between 0 and 5.", BRIGHTRED)
                Else
                    Call SendSetAccess(txtPlayerName.Text, txtValue.Text)
                End If
            End If
        End If
    Else
        Call AddText("You are not authorized to carry out that action.", BRIGHTRED)
    End If
End Sub

Private Sub btnSetNone_Click()
    Call SendData("weather" & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub btnSetRain_Click()
    Call SendData("weather" & SEP_CHAR & 1 & END_CHAR)
End Sub

Private Sub btnSetSnow_Click()
    Call SendData("weather" & SEP_CHAR & 2 & END_CHAR)
End Sub

Private Sub btnSetThunder_Click()
    Call SendData("weather" & SEP_CHAR & 3 & END_CHAR)
End Sub

Private Sub btnClose_Click()
    frmAdmin.Visible = False
End Sub

Private Sub CmdEditMain_Click()
    If Not TxtScriptName.Text = "" Then
        Call SendRequestEditMain(TxtScriptName.Text)
    Else
        Call AddText("You have to enter a script name to open it...", BRIGHTRED)
    End If
End Sub
