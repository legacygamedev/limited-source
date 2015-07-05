VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administration Options (Access level:"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlTime 
      Height          =   255
      Left            =   1560
      Max             =   30
      Min             =   1
      TabIndex        =   39
      Top             =   760
      Value           =   1
      Width           =   1815
   End
   Begin VB.CommandButton cmdMute 
      Caption         =   "(Un)Mute"
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
      Left            =   120
      TabIndex        =   11
      Top             =   760
      Width           =   1455
   End
   Begin VB.CommandButton cmdBan 
      Caption         =   "Ban"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGiveSelf 
      Caption         =   "Give/Take self PK"
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
      Left            =   1800
      TabIndex        =   38
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGiveTarget 
      Caption         =   "Give/Take target PK"
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
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   1695
   End
   Begin VB.HScrollBar scrlMap 
      Height          =   255
      Left            =   1800
      Max             =   100
      Min             =   1
      TabIndex        =   35
      Top             =   1920
      Value           =   1
      Width           =   2775
   End
   Begin VB.CommandButton cmdWarpTo 
      Caption         =   "Warp to map"
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
      Left            =   3000
      TabIndex        =   34
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdWarpTarget 
      Caption         =   "Warp target to me"
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
      Left            =   1560
      TabIndex        =   33
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdWarpMe 
      Caption         =   "Warp me to target"
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
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdRespawnMap 
      Caption         =   "Respawn map"
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
      Left            =   3000
      TabIndex        =   31
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDestroyBan 
      Caption         =   "Destroy ban list"
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
      Left            =   1560
      TabIndex        =   29
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
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
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdLevelTarget 
      Caption         =   "Level target"
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
      Left            =   3000
      TabIndex        =   26
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdLevelYourself 
      Caption         =   "Level yourself"
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
      Left            =   1560
      TabIndex        =   25
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheckInventory 
      Caption         =   "Check Inventory"
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
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   2640
      TabIndex        =   24
      ToolTipText     =   "Type in the name of the account for the check account button"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdCheckAccount 
      Caption         =   "Check Account"
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
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame FraEditors 
      Caption         =   "Editors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   4455
      Begin VB.CommandButton cmdAnim 
         Caption         =   "Anim"
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
         Left            =   3360
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSign 
         Caption         =   "Sign"
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
         Left            =   2280
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdShop 
         Caption         =   "Shop"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdMap 
         Caption         =   "Map"
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
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Item"
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
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSpell 
         Caption         =   "Spell"
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
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdNPC 
         Caption         =   "NPC"
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.HScrollBar scrlSpriteNum 
      Height          =   255
      Left            =   1800
      Max             =   100
      TabIndex        =   15
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdSetTarget 
      Caption         =   "Set target sprite"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetYour 
      Caption         =   "Set your sprite"
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
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick"
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
      Left            =   840
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtTargetPlayer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      ToolTipText     =   "Type in the name of the target player here (make sure they're online)"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   4455
      Begin VB.CheckBox chkHideAll 
         Caption         =   "Hide all text"
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
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkHideAnimations 
         Caption         =   "Hide animations"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkHideSprites 
         Caption         =   "Hide sprites"
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
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkItemPick 
         Caption         =   "Item pick-up macro"
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
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkRightClick 
         Caption         =   "Right click to warp"
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
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkDisplayCurrent 
         Caption         =   "Display current loc"
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
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkHideMap 
         Caption         =   "Hide map name"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label lblMinutes 
      Alignment       =   1  'Right Justify
      Caption         =   "1 Minutes"
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
      Left            =   3480
      TabIndex        =   40
      Top             =   780
      Width           =   1080
   End
   Begin VB.Label lblMapNumber 
      Caption         =   "Map Number: 1"
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
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblAccount 
      Caption         =   "Account:"
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
      Left            =   1800
      TabIndex        =   23
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblSpriteNumber 
      AutoSize        =   -1  'True
      Caption         =   "Sprite Number: 0"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Target Player:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDisplayCurrent_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    BLoc = chkDisplayCurrent.Value
    
End Sub

Private Sub cmdBan_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    If LenB(Trim$(txtTargetPlayer.Text)) < 1 Then
        AddText "The target player box is empty!", AlertColor
        Exit Sub
    End If
    
    SendBan Trim$(txtTargetPlayer.Text)
    
End Sub

Private Sub cmdCheckAccount_Click()
    SendACPAction CheckAccount
End Sub

Private Sub cmdCheckInventory_Click()
    SendACPAction CheckInventory
End Sub

Private Sub cmdDestroyBan_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Creator Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Creator & ")!", AlertColor
        Exit Sub
    End If
    
    SendBanDestroy
    
End Sub

Private Sub cmdGiveSelf_Click()
    SendACPAction GiveSelfPK
End Sub

Private Sub cmdGiveTarget_Click()
    SendACPAction GiveTargetPK
End Sub

Private Sub cmdHelp_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    AddText "Social Commands:", HelpColor
    AddText """msghere = Global Admin Message", HelpColor
    AddText "=msghere = Private Admin Message", HelpColor
    AddText "Available Commands: /acp, /admin, /loc, /editmap, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell, /editsign", HelpColor
    
End Sub

Private Sub cmdItem_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditItem
    
End Sub

Private Sub cmdKick_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    If LenB(Trim$(txtTargetPlayer.Text)) < 1 Then
        AddText "The target player box is empty!", AlertColor
        Exit Sub
    End If
    
    SendKick Trim$(txtTargetPlayer.Text)
    
End Sub

Private Sub cmdLevelTarget_Click()
    SendACPAction LevelTarget
End Sub

Private Sub cmdLevelYourself_Click()
    SendACPAction LevelSelf
End Sub

Private Sub cmdMap_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditMap
    
End Sub

Private Sub cmdMute_Click()
    SendACPAction MutePlayer
End Sub

Private Sub cmdNPC_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditNpc
    
End Sub

Private Sub cmdRespawnMap_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    SendMapRespawn
    
End Sub

Private Sub cmdSetTarget_Click()
    SendACPAction SetTargetSprite
End Sub

Private Sub cmdSetYour_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    SendSetSprite scrlSpriteNum.Value
    
End Sub

Private Sub cmdShop_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditShop
    
End Sub

Private Sub cmdSign_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditSign
    
End Sub

Private Sub cmdAnim_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditAnim
    
End Sub

Private Sub cmdSpell_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Developer & ")!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditSpell
    
End Sub
                            
Private Sub cmdWarpMe_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    If LenB(Trim$(txtTargetPlayer.Text)) < 1 Then
        AddText "You don't have a target player typed in!", AlertColor
        Exit Sub
    End If
    
    WarpMeTo Trim$(txtTargetPlayer.Text)
    
End Sub

Private Sub cmdWarpTarget_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    If LenB(Trim$(txtTargetPlayer.Text)) < 1 Then
        AddText "You don't have a target player typed in!", AlertColor
        Exit Sub
    End If
    
    WarpToMe Trim$(txtTargetPlayer.Text)
    
End Sub

Private Sub cmdWarpTo_Click()

    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    WarpTo scrlMap.Value
    
End Sub

Private Sub Form_Load()
    scrlSpriteNum.Max = TOTAL_SPRITES
    scrlMap.Max = MAX_MAPS
    Me.Caption = Me.Caption & " " & GetPlayerAccess(MyIndex) & ")"
End Sub

Private Sub scrlMap_Change()
    lblMapNumber.Caption = "Map Number: " & scrlMap.Value
End Sub

Private Sub scrlMap_Scroll()
    scrlMap_Change
End Sub

Private Sub scrlSpriteNum_Change()
    lblSpriteNumber.Caption = "Sprite number: " & scrlSpriteNum.Value
End Sub

Private Sub scrlSpriteNum_Scroll()
    scrlSpriteNum_Change
End Sub

Private Sub scrlTime_Change()
    lblMinutes.Caption = scrlTime.Value & " Minutes"
End Sub

Private Sub scrlTime_Scroll()
    scrlTime_Change
End Sub
