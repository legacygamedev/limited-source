VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Settings"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPK 
      Caption         =   "Player Killing - (PVP)"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CheckBox chkNPCCORPSE 
      Caption         =   "Npc Corpses"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CheckBox chkPLAYERCORPSE 
      Caption         =   "Player Corpses"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtSpriteY 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtSpriteX 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Classes"
      Height          =   975
      Left            =   2520
      TabIndex        =   20
      Top             =   3000
      Width           =   1455
      Begin VB.CommandButton Command29 
         Caption         =   "Reload"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Edit"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scripts"
      Height          =   1335
      Left            =   2520
      TabIndex        =   15
      Top             =   1680
      Width           =   1455
      Begin VB.CommandButton Command25 
         Caption         =   "Reload"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Turn On"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Turn Off"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Edit"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Text            =   "Checked = ON | UN-Checked = OFF"
      Top             =   120
      Width           =   4335
   End
   Begin VB.CheckBox chkLanguage 
      Caption         =   "Language Filter"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox chkMovement 
      Caption         =   "Movement Tiredness"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox chkKickIdlePlayers 
      Caption         =   "Kick Idle Players (10 Mins)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CheckBox chkMPRegen 
      Caption         =   "MP Regeneration"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CheckBox chkSPRegen 
      Caption         =   "SP Regeneration"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CheckBox chkHPRegen 
      Caption         =   "HP Regeneration"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox chkPaperdoll 
      Caption         =   "Paperdoll"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chkEXPLOSS 
      Caption         =   "Experience On Death"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit Program"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Settings && Run Server"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Enter The (Y ) for Players Sprite Size"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Enter The ( X ) for Players Sprite Size"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label lblDate 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Please Enter Your Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

If txtName.text <> "" Then

If txtSpriteX.text = "" Then
MsgBox "Enter an X for your spritesize"
Exit Sub
End If

If txtSpriteY.text = "" Then
MsgBox "Enter a Y for your spritesize"
Exit Sub
End If

If frmSettings.chkHPRegen.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 0
End If

If frmSettings.chkPK.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_KILLING", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_KILLING", 0
End If

If frmSettings.chkSPRegen.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 0
End If

If frmSettings.chkMPRegen.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 0
End If

If frmSettings.chkEXPLOSS.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "DEATHEXPLOSS", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "DEATHEXPLOSS", 0
End If

If frmSettings.chkKickIdlePlayers.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "KICKIDLEPLAYERS", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "KICKIDLEPLAYERS", 0
End If

If frmSettings.chkLanguage.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "LANGUAGEFILTER", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "LANGUAGEFILTER", 0
End If

If frmSettings.chkPaperdoll.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PAPERDOLL", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PAPERDOLL", 0
End If

If frmSettings.chkMovement.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MOVEMENT_TIREDNESS", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MOVEMENT_TIREDNESS", 0
End If

If frmSettings.chkPLAYERCORPSE.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_CORPSES", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_CORPSES", 0
End If

If frmSettings.chkNPCCORPSE.Value = Checked Then
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "NPC_CORPSES", 1
Else
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "NPC_CORPSES", 0
End If

SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_SIZE_X", "" & txtSpriteX.text
SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PLAYER_SIZE_y", "" & txtSpriteY.text

Reg(1).Name = Trim(txtName.text)
Reg(1).regdate = Date
Reg(1).regtime = Time
Call SaveREG(1)

frmSettings.Visible = False

If ServerOnline = 0 Then
Call ServerLoop
End If

PK = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PK"))
PLAYER_CORPSES = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PLAYER_CORPSES"))
NPC_CORPSES = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "NPC_CORPSES"))
SIZE_X = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PLAYER_SIZE_X"))
SIZE_Y = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PLAYER_SIZE_Y"))
HPRegen = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRgen"))
SPRegen = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRgen"))
MPRegen = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRgen"))
SCRIPTING = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SCRIPTING"))
MAX_ELEMENTS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS"))
PAPERDOLL = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PAPERDOLL"))
SPRITESIZE = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRITESIZE"))
MOVEMENT_TIREDNESS = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MOVEMENT_TIREDNESS"))
'POINTS_PER_LEVEL = GetVar(App.Path & "\Data.ini", "CONFIG", "PointsPerLevel")
LANGUAGEFILTER = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "LANGUAGEFILTER"))
DEATHEXPLOSS = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "DEATHEXPLOSS"))
KICKIDLEPLAYERS = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "KICKIDLEPLAYERS"))
Else
MsgBox "Please Enter your Name !"
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Call DestroyServer
End Sub

Private Sub Command25_Click()

    If SCRIPTING = 1 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\main\Scripts\Main.txt", "main\Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
    End If
End Sub

Private Sub Command26_Click()

    If SCRIPTING = 0 Then
        SCRIPTING = 1
        PutVar App.Path & "\Data.ini", "CONFIG", "SCRIPTING", 1

        If SCRIPTING = 1 Then
            Set MyScript = New clsSadScript
            Set clsScriptCommands = New clsCommands
            MyScript.ReadInCode App.Path & "\main\Scripts\Main.txt", "main\Scripts\Main.txt", MyScript.SControl, False
            MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        End If
    End If
End Sub

Private Sub Command27_Click()

    If SCRIPTING = 1 Then
        SCRIPTING = 0
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SCRIPTING", 0

        If SCRIPTING = 0 Then
            Set MyScript = Nothing
            Set clsScriptCommands = Nothing
        End If
    End If
End Sub

Private Sub Command28_Click()
AFileName = "main\Scripts/Main.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command29_Click()
Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
End Sub

Private Sub Command30_Click()
AFileName = "main\Classes\Info.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Form_Load()
If Trim(Reg(1).Name) <> "" Then
txtName.Visible = False
Label2.Visible = False
lblName.Caption = Trim("Registered To: " & Reg(1).Name)
lblDate.Caption = Trim(Reg(1).regdate & " - " & Reg(1).regtime)
txtName.text = Trim(Reg(1).Name)
End If
frmSettings.chkEXPLOSS.Value = DEATHEXPLOSS
frmSettings.chkPaperdoll.Value = PAPERDOLL
frmSettings.chkLanguage.Value = LANGUAGEFILTER
frmSettings.chkKickIdlePlayers.Value = KICKIDLEPLAYERS
frmSettings.chkMovement.Value = MOVEMENT_TIREDNESS
frmSettings.chkHPRegen.Value = HPRegen
frmSettings.chkSPRegen.Value = SPRegen
frmSettings.chkMPRegen.Value = MPRegen
frmSettings.txtSpriteX.text = SIZE_X
frmSettings.txtSpriteY.text = SIZE_Y
frmSettings.chkPLAYERCORPSE.Value = PLAYER_CORPSES
frmSettings.chkNPCCORPSE.Value = NPC_CORPSES
frmSettings.chkPK.Value = PK
End Sub
