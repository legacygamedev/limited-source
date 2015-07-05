VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCClient 
   BorderStyle     =   0  'None
   Caption         =   "Cerberus"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   766
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1022
      TabIndex        =   3
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox picRightClickMenu 
         Height          =   2895
         Left            =   12600
         ScaleHeight     =   189
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   117
         TabIndex        =   4
         Top             =   5040
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CommandButton cmdPlayerChat 
            Caption         =   "Chat"
            Height          =   375
            Left            =   120
            TabIndex        =   170
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmdPlayerQuests 
            Caption         =   "Quests"
            Height          =   375
            Left            =   120
            TabIndex        =   146
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton cmdPlayerSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdPlayerSkills 
            Caption         =   "Skills"
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdPlayerStats 
            Caption         =   "Stats"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdInventory 
            Caption         =   "Inventory"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton cmdQuit 
            Caption         =   "Quit"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Timer tmrRightClick 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   1320
            Top             =   1680
         End
      End
      Begin VB.PictureBox picPlayerChat 
         Height          =   2175
         Left            =   4440
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   480
         TabIndex        =   165
         Top             =   7560
         Visible         =   0   'False
         Width           =   7260
         Begin VB.CheckBox chkPlayerChatPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   169
            Top             =   90
            Width           =   615
         End
         Begin RichTextLib.RichTextBox txtChat 
            Height          =   1335
            Left            =   120
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2355
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmCClient.frx":0000
         End
         Begin VB.Timer tmrPlayerChat 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   840
            Top             =   120
         End
         Begin VB.Label lblPlayerChatCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   6840
            TabIndex        =   168
            Top             =   60
            Width           =   165
         End
         Begin VB.Label lblChat 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   1800
            Width           =   6975
         End
      End
      Begin VB.PictureBox picNpcQuests 
         Height          =   4215
         Left            =   6000
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   271
         TabIndex        =   147
         Top             =   3480
         Visible         =   0   'False
         Width           =   4125
         Begin VB.CommandButton cmdAbandonQuest 
            Caption         =   "Abandon"
            Height          =   375
            Left            =   1440
            TabIndex        =   163
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtNpcQuestDesc 
            Height          =   975
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   153
            Top             =   1560
            Width           =   3615
         End
         Begin VB.ListBox lstNpcQuests 
            Height          =   1035
            ItemData        =   "frmCClient.frx":0082
            Left            =   240
            List            =   "frmCClient.frx":0084
            TabIndex        =   152
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdCancelQuest 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   150
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdCompleteQuest 
            Caption         =   "Complete"
            Height          =   375
            Left            =   1440
            TabIndex        =   149
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdAcceptQuest 
            Caption         =   "Accept"
            Height          =   375
            Left            =   120
            TabIndex        =   148
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label lblNpcQuestMax 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3360
            TabIndex        =   162
            Top             =   3000
            Width           =   90
         End
         Begin VB.Label lblNpcQuestMin 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1920
            TabIndex        =   161
            Top             =   3000
            Width           =   90
         End
         Begin VB.Label Label20 
            Caption         =   "Maximum :"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   2400
            TabIndex        =   160
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Minimum :"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   1080
            TabIndex        =   159
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Level :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   158
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label lblNpcQuestReward 
            Height          =   255
            Left            =   1560
            TabIndex        =   157
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label Label17 
            Caption         =   "Reward :"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   480
            TabIndex        =   156
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label lblNpcQuestClass 
            Height          =   255
            Left            =   1800
            TabIndex        =   155
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label16 
            Caption         =   "Class Required :"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   240
            TabIndex        =   154
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblQuestNpc 
            Alignment       =   2  'Center
            Caption         =   "Npc Name"
            Height          =   255
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   4095
         End
      End
      Begin VB.PictureBox picPlayerQuests 
         AutoRedraw      =   -1  'True
         Height          =   2775
         Left            =   6120
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   136
         Top             =   360
         Visible         =   0   'False
         Width           =   3900
         Begin VB.CheckBox chkPlayerQuestPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   120
            Width           =   615
         End
         Begin VB.HScrollBar scrlPlayerQuestNum 
            Height          =   255
            Left            =   1080
            Max             =   5
            Min             =   1
            TabIndex        =   138
            Top             =   480
            Value           =   1
            Width           =   2175
         End
         Begin VB.Timer tmrPlayerQuests 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   2400
            Top             =   120
         End
         Begin VB.TextBox txtPlayerQuestDesc 
            Height          =   975
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   141
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label lblPlayerQuestCount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "-- / --"
            Height          =   195
            Left            =   3225
            TabIndex        =   164
            Top             =   2040
            Width           =   345
         End
         Begin VB.Label lblPlayerQuestQuit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3540
            TabIndex        =   145
            Top             =   120
            Width           =   165
         End
         Begin VB.Label lblPlayerQuestReward 
            Alignment       =   2  'Center
            Caption         =   "Item"
            Height          =   255
            Left            =   0
            TabIndex        =   143
            Top             =   2400
            Width           =   3855
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Reward"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   0
            TabIndex        =   142
            Top             =   2160
            Width           =   3855
         End
         Begin VB.Label lblPlayerQuestName 
            Alignment       =   2  'Center
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label lblPlayerQuestNum 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Left            =   3360
            TabIndex        =   139
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label14 
            Caption         =   "Quest No."
            Height          =   255
            Left            =   240
            TabIndex        =   137
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox picPlayerSkills 
         Height          =   1230
         Left            =   9180
         ScaleHeight     =   78
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   397
         TabIndex        =   97
         Top             =   10080
         Visible         =   0   'False
         Width           =   6015
         Begin VB.PictureBox picSkillDesc 
            Height          =   1170
            Left            =   3600
            ScaleHeight     =   74
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   77
            TabIndex        =   116
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
            Begin VB.Label lblSkillDescPer 
               Alignment       =   2  'Center
               Caption         =   "per Level"
               ForeColor       =   &H00C000C0&
               Height          =   225
               Left            =   0
               TabIndex        =   121
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label lblSkillDescMod 
               Alignment       =   2  'Center
               Caption         =   "Mod"
               ForeColor       =   &H00C000C0&
               Height          =   225
               Left            =   0
               TabIndex        =   120
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label lblSkillDescExp 
               Alignment       =   2  'Center
               Caption         =   "Exp"
               Height          =   225
               Left            =   0
               TabIndex        =   119
               Top             =   450
               Width           =   1215
            End
            Begin VB.Label lblSkillDescLevel 
               Alignment       =   2  'Center
               Caption         =   "Level"
               Height          =   225
               Left            =   0
               TabIndex        =   118
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label lblSkillDescName 
               Alignment       =   2  'Center
               Caption         =   "Name"
               ForeColor       =   &H80000002&
               Height          =   225
               Left            =   0
               TabIndex        =   117
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.ListBox lstPlayerSkills 
            Height          =   1620
            ItemData        =   "frmCClient.frx":0086
            Left            =   120
            List            =   "frmCClient.frx":0088
            TabIndex        =   112
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Timer tmrPlayerSkills 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3480
            Top             =   1320
         End
         Begin VB.PictureBox picSkillSprite 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4080
            ScaleHeight     =   33
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   33
            TabIndex        =   111
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox picPlayerSkillsBack 
            BackColor       =   &H8000000C&
            Height          =   735
            Left            =   120
            ScaleHeight     =   45
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   379
            TabIndex        =   99
            Top             =   360
            Width           =   5745
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   5040
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   110
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   4500
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   109
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   3960
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   108
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   3420
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   107
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   2880
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   106
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   2280
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   105
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1740
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   104
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1200
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   103
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   660
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   102
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSkill 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   101
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.CheckBox chkPlayerSkillsPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   75
            Width           =   615
         End
         Begin VB.Label lblPlayerSkillsCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   5700
            TabIndex        =   100
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox picPlayerStats 
         Height          =   4815
         Left            =   12780
         ScaleHeight     =   317
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
         Begin VB.PictureBox picPlayerArrows 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   134
            Top             =   3240
            Width           =   495
            Begin VB.Label lblArrows 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   90
               TabIndex        =   135
               Top             =   285
               Width           =   375
            End
         End
         Begin VB.PictureBox picPlayerRing 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1680
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   115
            Top             =   3480
            Width           =   495
         End
         Begin VB.PictureBox picPlayerAmulet 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1680
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   114
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox picPlayerHelmet 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   3600
            Width           =   495
         End
         Begin VB.PictureBox picPlayerShield 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1560
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   4200
            Width           =   495
         End
         Begin VB.PictureBox picPlayerArmour 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   4200
            Width           =   495
         End
         Begin VB.PictureBox picPlayerWeapon 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   360
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   4200
            Width           =   495
         End
         Begin VB.CheckBox chkPlayerStatsPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   120
            Width           =   615
         End
         Begin VB.Timer tmrPlayerStats 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   1560
            Top             =   480
         End
         Begin VB.Label lblDEXTERITY 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   133
            Top             =   2280
            Width           =   90
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "DEXTERITY"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1080
            TabIndex        =   131
            Top             =   2880
            Width           =   90
         End
         Begin VB.Label lblLEVEL 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1080
            TabIndex        =   130
            Top             =   2640
            Width           =   90
         End
         Begin VB.Label Label12 
            Caption         =   "EXP"
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "LEVEL"
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblSTRENGTH 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   20
            Top             =   1320
            Width           =   90
         End
         Begin VB.Label lblMAGIC 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   23
            Top             =   2040
            Width           =   90
         End
         Begin VB.Label lblSPEED 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   22
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label lblDEFENCE 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label Label7 
            Caption         =   "MAGIC"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "SPEED"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "DEFENCE"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "STRENGTH"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblPlayerStatsQuit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2040
            TabIndex        =   14
            Top             =   90
            Width           =   165
         End
         Begin VB.Label Label3 
            Caption         =   "SP"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "MP"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "HP"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lblSP 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   840
            TabIndex        =   9
            Top             =   960
            Width           =   90
         End
         Begin VB.Label lblMP 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   840
            TabIndex        =   8
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblHP 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   840
            TabIndex        =   7
            Top             =   480
            Width           =   90
         End
      End
      Begin VB.PictureBox picPlayerInventory 
         Height          =   6570
         Left            =   120
         ScaleHeight     =   434
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   139
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   2145
         Begin VB.PictureBox picItemDesc 
            Height          =   1935
            Left            =   360
            ScaleHeight     =   125
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   90
            TabIndex        =   88
            Top             =   0
            Width           =   1410
            Begin VB.Label lblItemDescSpell 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "ItemDescSpell"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   96
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Spell Held"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   0
               TabIndex        =   92
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lblItemDescMod 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "ItemDescMod"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   95
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "Attribute Modified"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   0
               TabIndex        =   91
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblItemDescReq 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Caption         =   "ItemDescReq"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   94
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Requirement"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   0
               TabIndex        =   93
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblItemDescDur 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemDescDur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   90
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblItemDescName 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemDescName"
               ForeColor       =   &H80000002&
               Height          =   255
               Left            =   0
               TabIndex        =   89
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.PictureBox picInventoryItems 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2400
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox picPlayerInventoryBack 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6060
            Left            =   120
            ScaleHeight     =   404
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   124
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   360
            Width           =   1860
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   29
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   5460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   28
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   5460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   27
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   5460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   26
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   4860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   25
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   4860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   24
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   4860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   23
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   4260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   22
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   4260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   21
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   4260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   20
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   3660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   19
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   3660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   18
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   3660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   17
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   3060
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   16
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   3060
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   15
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   3060
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   14
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   2460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   13
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   2460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   12
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   2460
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   11
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   10
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   1860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   9
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   1860
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   8
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   1260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   7
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   1260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   6
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1260
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   5
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   4
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   3
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   660
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   2
               Left            =   1260
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   60
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   1
               Left            =   660
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   60
               Width           =   540
            End
            Begin VB.PictureBox picPlayerInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   0
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   60
               Width           =   540
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   6
               Left            =   1230
               Top             =   5430
               Width           =   600
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   5
               Left            =   1230
               Top             =   4830
               Width           =   600
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   4
               Left            =   1230
               Top             =   4230
               Width           =   600
            End
            Begin VB.Shape shpSelectedItem 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   585
               Left            =   45
               Top             =   45
               Width           =   585
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   3
               Left            =   1230
               Top             =   3630
               Width           =   600
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   2
               Left            =   1230
               Top             =   3030
               Width           =   600
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   1
               Left            =   1230
               Top             =   2430
               Width           =   600
            End
            Begin VB.Shape shpEquiped 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   600
               Index           =   0
               Left            =   1230
               Top             =   1830
               Width           =   600
            End
         End
         Begin VB.CheckBox chkPlayerInventoryPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   60
            Width           =   615
         End
         Begin VB.Timer tmrPlayerInventory 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3120
            Top             =   2160
         End
         Begin VB.ListBox lstPlayerInventory 
            Height          =   1620
            Left            =   2400
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblPlayerInventoryQuit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1800
            TabIndex        =   2
            Top             =   60
            Width           =   165
         End
      End
      Begin VB.PictureBox picShopTrade 
         Height          =   4095
         Left            =   4320
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   501
         TabIndex        =   67
         Top             =   3240
         Visible         =   0   'False
         Width           =   7575
         Begin VB.PictureBox picFixItems 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3015
            Left            =   120
            ScaleHeight     =   3015
            ScaleWidth      =   7335
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   7335
            Begin VB.CommandButton cmdItemFix 
               Caption         =   "Fix"
               Height          =   375
               Left            =   2760
               TabIndex        =   75
               Top             =   2280
               Width           =   1935
            End
            Begin VB.ListBox lstFixItem 
               Height          =   1815
               ItemData        =   "frmCClient.frx":008A
               Left            =   120
               List            =   "frmCClient.frx":008C
               TabIndex        =   74
               Top             =   120
               Width           =   7095
            End
         End
         Begin VB.CommandButton cmdShopCancel 
            Caption         =   "Close Shop"
            Height          =   375
            Left            =   2880
            TabIndex        =   72
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CommandButton cmdShopDeal 
            Caption         =   "MakeTrade"
            Height          =   375
            Left            =   2880
            TabIndex        =   71
            Top             =   3000
            Width           =   1935
         End
         Begin VB.CommandButton cmdShopFixItems 
            Caption         =   "Fix Items"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   70
            Top             =   2520
            Width           =   1935
         End
         Begin VB.ListBox lstShopTrade 
            Height          =   1815
            ItemData        =   "frmCClient.frx":008E
            Left            =   240
            List            =   "frmCClient.frx":0090
            TabIndex        =   69
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label lblShopName 
            Alignment       =   2  'Center
            Caption         =   "Shop Name"
            Height          =   255
            Left            =   2040
            TabIndex        =   68
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.PictureBox picPlayerSpells 
         Height          =   1230
         Left            =   120
         ScaleHeight     =   78
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   397
         TabIndex        =   62
         Top             =   10020
         Visible         =   0   'False
         Width           =   6015
         Begin VB.PictureBox picSpellDesc 
            Height          =   1170
            Left            =   3600
            ScaleHeight     =   74
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   77
            TabIndex        =   122
            Top             =   0
            Width           =   1215
            Begin VB.Label lblSpellDescMod 
               Alignment       =   2  'Center
               Caption         =   "Mod"
               ForeColor       =   &H00C000C0&
               Height          =   225
               Left            =   0
               TabIndex        =   126
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label lblSpellDescName 
               Alignment       =   2  'Center
               Caption         =   "Name"
               ForeColor       =   &H80000002&
               Height          =   225
               Left            =   0
               TabIndex        =   123
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label lblSpellDescPer 
               Alignment       =   2  'Center
               Caption         =   "per Level"
               ForeColor       =   &H00C000C0&
               Height          =   225
               Left            =   0
               TabIndex        =   127
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label lblSpellDescExp 
               Alignment       =   2  'Center
               Caption         =   "Exp"
               Height          =   225
               Left            =   0
               TabIndex        =   125
               Top             =   450
               Width           =   1215
            End
            Begin VB.Label lblSpellDescLevel 
               Alignment       =   2  'Center
               Caption         =   "Level"
               Height          =   225
               Left            =   0
               TabIndex        =   124
               Top             =   210
               Width           =   1215
            End
         End
         Begin VB.PictureBox picPlayerSpellsBack 
            BackColor       =   &H8000000C&
            Height          =   735
            Left            =   120
            ScaleHeight     =   45
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   379
            TabIndex        =   77
            Top             =   360
            Width           =   5745
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   5040
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   87
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   4500
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   86
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   3960
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   85
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   3420
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   84
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   2880
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   83
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   2280
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   82
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1740
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   81
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1200
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   80
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   660
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   79
               Top             =   90
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   78
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.PictureBox picSpellSprite 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4080
            ScaleHeight     =   33
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   33
            TabIndex        =   76
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Timer tmrPlayerSpells 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3480
            Top             =   1320
         End
         Begin VB.ListBox lstPlayerSpells 
            Height          =   2010
            ItemData        =   "frmCClient.frx":0092
            Left            =   120
            List            =   "frmCClient.frx":0094
            TabIndex        =   66
            Top             =   1320
            Width           =   3135
         End
         Begin VB.CheckBox chkPlayerSpellsPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lblPlayerSpellsCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   5640
            TabIndex        =   64
            Top             =   75
            Width           =   165
         End
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   14760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCClient"
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

Dim MouseXOffset, MouseYOffset As Integer

Private Sub Form_Load()
    'If picInventoryItems.Picture = LoadPicture() Then
        picInventoryItems.Picture = LoadPicture(App.Path & GFX_PATH & "Items" & GFX_EXT)
    'End If
    'If picSpellSprite.Picture = LoadPicture() Then
        picSpellSprite.Picture = LoadPicture(App.Path & GFX_PATH & "Spells" & GFX_EXT)
    'End If
    'If picSkillSprite.Picture = LoadPicture() Then
        picSkillSprite.Picture = LoadPicture(App.Path & GFX_PATH & "Skills" & GFX_EXT)
    'End If
    
    If Prefs.Inventory = 0 Then
        picPlayerInventory.Visible = False
        chkPlayerInventoryPin.Value = Unchecked
    Else
        picPlayerInventory.Visible = True
        chkPlayerInventoryPin.Value = Checked
    End If
    If Prefs.Stats = 0 Then
        picPlayerStats.Visible = False
        chkPlayerStatsPin.Value = Unchecked
    Else
        picPlayerStats.Visible = True
        chkPlayerStatsPin.Value = Checked
    End If
    If Prefs.Skills = 0 Then
        picPlayerSkills.Visible = False
        chkPlayerSkillsPin.Value = Unchecked
    Else
        picPlayerSkills.Visible = True
        chkPlayerSkillsPin.Value = Checked
    End If
    If Prefs.Spells = 0 Then
        picPlayerSpells.Visible = False
        chkPlayerSpellsPin.Value = Unchecked
    Else
        picPlayerSpells.Visible = True
        chkPlayerSpellsPin.Value = Checked
    End If
    If Prefs.Quests = 0 Then
        picPlayerQuests.Visible = False
        chkPlayerQuestPin.Value = Unchecked
    Else
        picPlayerQuests.Visible = True
        chkPlayerQuestPin.Value = Checked
    End If
    If Prefs.Chat = 0 Then
        picPlayerChat.Visible = False
        chkPlayerChatPin.Value = Unchecked
    Else
        picPlayerChat.Visible = True
        chkPlayerChatPin.Value = Checked
    End If
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub picScreen_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
End Sub

Private Sub txtChat_GotFocus()
    frmCClient.picScreen.SetFocus
End Sub

Private Sub picScreen_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x - MouseXOffset, y - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Int(x / 32) > (MAX_MAPX - 4) Or Int(y / 32) > (MAX_MAPY - 5) Then
            frmCClient.picRightClickMenu.top = y - 120
            frmCClient.picRightClickMenu.Left = x - 100
            frmCClient.picRightClickMenu.Visible = True
            frmCClient.tmrRightClick.Interval = 3000
            frmCClient.tmrRightClick.Enabled = True
        Else
            frmCClient.picRightClickMenu.top = y - 20
            frmCClient.picRightClickMenu.Left = x - 20
            frmCClient.picRightClickMenu.Visible = True
            frmCClient.tmrRightClick.Interval = 3000
            frmCClient.tmrRightClick.Enabled = True
        End If
    Else
        If Button = 1 Then
            picRightClickMenu.Visible = False
        End If
    End If
    'Call EditorMouseDown(Button, Shift, x, y)
    Call PlayerSearch(Button, Shift, x, y)
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picItemDesc.Visible = False
    picSkillDesc.Visible = False
    picSpellDesc.Visible = False
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub picScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    picRightClickMenu.Visible = False
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub picScreen_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckInput(0, KeyCode, Shift)
End Sub

' ********************
' * Right Click Menu *
' ********************

Private Sub cmdPlayerStats_Click()
    picRightClickMenu.Visible = False
    chkPlayerStatsPin.Value = Unchecked
    picPlayerStats.Visible = True
    tmrPlayerStats.Enabled = True
End Sub

Private Sub cmdInventory_Click()
    Call UpdateInventory
    Call UpdateVisInventory
    picRightClickMenu.Visible = False
    chkPlayerInventoryPin.Value = Unchecked
    picPlayerInventory.Visible = True
    tmrPlayerInventory.Enabled = True
End Sub

Private Sub cmdPlayerSpells_Click()
    'Call SendData("spells" & SEP_CHAR & END_CHAR)
    picRightClickMenu.Visible = False
    chkPlayerSpellsPin.Value = Unchecked
    picPlayerSpells.Visible = True
    'Call UpdateVisSpells
    tmrPlayerSpells.Enabled = True
End Sub

Private Sub cmdPlayerSkills_Click()
    'Call SendData("skills" & SEP_CHAR & END_CHAR)
    picRightClickMenu.Visible = False
    chkPlayerSkillsPin.Value = Unchecked
    picPlayerSkills.Visible = True
    'Call UpdateVisSkills
    tmrPlayerSkills.Enabled = True
End Sub

Private Sub cmdPlayerQuests_Click()
    picRightClickMenu.Visible = False
    chkPlayerQuestPin.Value = Unchecked
    picPlayerQuests.Visible = True
    tmrPlayerQuests.Enabled = True
End Sub

Private Sub cmdPlayerChat_Click()
    picRightClickMenu.Visible = False
    chkPlayerChatPin.Value = Unchecked
    picPlayerChat.Visible = True
    tmrPlayerChat.Enabled = True
End Sub

Private Sub cmdQuit_Click()
    Call GameDestroy
End Sub
Private Sub tmrRightClick_Timer()
    picRightClickMenu.Visible = False
    tmrRightClick.Enabled = False
End Sub

' ****************
' * Player Stats *
' ****************

Private Sub tmrPlayerStats_Timer()
    picPlayerStats.Visible = False
    tmrPlayerStats.Enabled = False
End Sub

Private Sub lblPlayerStatsQuit_Click()
    picPlayerStats.Visible = False
    tmrPlayerStats.Enabled = False
    chkPlayerStatsPin.Value = Unchecked
End Sub

Private Sub chkPlayerStatsPin_Click()
    If chkPlayerStatsPin.Value = Checked Then
        tmrPlayerStats.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerStats.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub picPlayerStats_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerStats.Drag vbBeginDrag
End Sub

Private Sub picPlayerStats_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerStats.Left - MouseXOffset, y + picPlayerStats.top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub picPlayerHelmet_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerWeapon_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerArmour_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerShield_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerAmulet_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerRing_Click()
    picScreen.SetFocus
End Sub

Private Sub picPlayerArrow_Click()
    picScreen.SetFocus
End Sub

' ********************
' * Player Inventory *
' ********************

Private Sub tmrPlayerInventory_Timer()
    picPlayerInventory.Visible = False
    tmrPlayerInventory.Enabled = False
End Sub

Private Sub lblPlayerInventoryQuit_Click()
    picPlayerInventory.Visible = False
    tmrPlayerInventory.Enabled = False
    chkPlayerInventoryPin.Value = Unchecked
End Sub

Private Sub chkPlayerInventoryPin_Click()
    If chkPlayerInventoryPin.Value = Checked Then
        tmrPlayerInventory.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerInventory.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub picPlayerInventoryBack_Click()
    picScreen.SetFocus
    'picPlayerInventory.Visible = True
End Sub

Private Sub picPlayerInv_DblClick(Index As Integer)
Dim d As Long

If GetPlayerInvItemNum(MyIndex, lstPlayerInventory.ListIndex + 1) = ITEM_TYPE_NONE Then Exit Sub

Call SendUseItem(lstPlayerInventory.ListIndex + 1)

'For d = 1 To MAX_INV
    'If GetPlayerInvItemNum(MyIndex, lstPlayerInventory.ListIndex + d) > 0 Then
        'If Item(GetPlayerInvItemNum(MyIndex, lstPlayerInventory.ListIndex + d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            'picPlayerInv(d - 1).Picture = LoadPicture()
        'End If
    'End If
'Next d
End Sub

Private Sub picPlayerInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Value As Long
Dim InvNum As Long
lstPlayerInventory.ListIndex = Index

    If Button = 1 Then
        Call UpdateVisInventory
    ElseIf Button = 2 Then
        If GetPlayerInvItemNum(MyIndex, lstPlayerInventory.ListIndex + 1) = ITEM_TYPE_NONE Then Exit Sub
        
        InvNum = frmCClient.lstPlayerInventory.ListIndex + 1
    
        If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Show them the drop dialog
                frmDrop.top = y - 20
                frmDrop.Left = x - 20
                'frmCClient.picScreen.SetFocus
                frmDrop.Show vbModal
            Else
                Call SendDropItem(frmCClient.lstPlayerInventory.ListIndex + 1, 0)
            End If
        End If
       
        picPlayerInv(InvNum - 1).Picture = LoadPicture()
        Call UpdateVisInventory
    End If
End Sub

Private Sub picItemDesc_GotFocus()
    picScreen.SetFocus
End Sub

Private Sub picPlayerInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim InvNum As Long
Dim ItemNum As Long
Dim ItemType As Long
Dim ItemDur As Long

    InvNum = Index + 1
    
    If GetPlayerInvItemNum(MyIndex, InvNum) <= 0 Then
        picItemDesc.Visible = False
        Exit Sub
    Else
        ItemNum = GetPlayerInvItemNum(MyIndex, InvNum)
        ItemType = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type
        ItemDur = GetPlayerInvItemDur(MyIndex, InvNum)
    End If
    
    lblItemDescName.Caption = Trim(Item(ItemNum).Name)
    Select Case ItemType
        Case ITEM_TYPE_NONE
            Exit Sub
        
        Case ITEM_TYPE_WEAPON
            picItemDesc.Height = 65
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 48
            If Item(ItemNum).Data3 = WEAPON_SUBTYPE_BOW Then
                lblItemDescReq.Caption = "Strength: " & Item(ItemNum).Data2
            ElseIf (Item(ItemNum).Data3 >= WEAPON_SUBTYPE_WAND) And (Item(ItemNum).Data3 <= WEAPON_SUBTYPE_STAFF) Then
                lblItemDescReq.Caption = "Magic: " & Item(ItemNum).Data2
            Else
                lblItemDescReq.Caption = "Strength: " & Item(ItemNum).Data2
            End If
            
        Case ITEM_TYPE_ARMOR
            picItemDesc.Height = 65
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 48
            lblItemDescReq.Caption = "Defence: " & Item(ItemNum).Data2
            
        Case ITEM_TYPE_HELMET
            picItemDesc.Height = 65
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 48
            lblItemDescReq.Caption = "Defence: " & Item(ItemNum).Data2
            
        Case ITEM_TYPE_SHIELD
            picItemDesc.Height = 65
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 48
            lblItemDescReq.Caption = "Defence: " & Item(ItemNum).Data2
            
        Case ITEM_TYPE_TOOL
            picItemDesc.Height = 65
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 48
            lblItemDescReq.Caption = "Strength: " & Item(ItemNum).Data2
            
        Case ITEM_TYPE_POTIONADDHP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "HP + " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTIONADDMP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "MP + " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTIONADDSP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "SP + " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTIONSUBHP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "HP - " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTIONSUBMP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "MP - " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTIONSUBSP
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescMod.top = 32
            lblItemDescMod.Caption = "SP - " & Item(ItemNum).Data1
            
        Case ITEM_TYPE_KEY
            lblItemDescDur.Caption = ""
            Label9.top = 64
            Label10.top = 96
            picItemDesc.Height = 33
            lblItemDescDur.Caption = "Key"
            
        Case ITEM_TYPE_CURRENCY
            lblItemDescDur.Caption = ""
            Label9.top = 64
            Label10.top = 96
            picItemDesc.Height = 33
            lblItemDescDur.Caption = "Carrying: " & GetPlayerInvItemValue(MyIndex, InvNum)
            
        Case ITEM_TYPE_SPELL
            picItemDesc.Height = 65
            lblItemDescDur.Caption = ""
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 32
            Label10.Caption = "Spell Held"
            lblItemDescSpell.top = 48
            lblItemDescSpell.Caption = Trim(Spell(Item(ItemNum).Data1).Name)
            
        Case ITEM_TYPE_AMULET
            picItemDesc.Height = 49
            lblItemDescDur.Caption = ""
            Label9.top = 16
            Label10.top = 96
            lblItemDescMod.top = 32
            Select Case Item(ItemNum).Data1
                Case CHARM_TYPE_ADDHP
                    lblItemDescMod.Caption = "HP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDMP
                    lblItemDescMod.Caption = "MP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSP
                    lblItemDescMod.Caption = "SP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSTR
                    lblItemDescMod.Caption = "Strength +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDDEF
                    lblItemDescMod.Caption = "Defence +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDMAGI
                    lblItemDescMod.Caption = "Magic +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSPEED
                    lblItemDescMod.Caption = "Speed +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDDEX
                    lblItemDescMod.Caption = "Dexterity +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDCRIT
                    lblItemDescMod.Caption = "Critical +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDDROP
                    lblItemDescMod.Caption = "Drop +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDBLOCK
                    lblItemDescMod.Caption = "Block +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDACCU
                    lblItemDescMod.Caption = "Accuracy +" & Item(ItemNum).Data2 & "%"
            End Select
            
        Case ITEM_TYPE_RING
                picItemDesc.Height = 49
                lblItemDescDur.Caption = ""
                Label9.top = 16
                Label10.top = 96
                lblItemDescMod.top = 32
            Select Case Item(ItemNum).Data1
                Case CHARM_TYPE_ADDHP
                    lblItemDescMod.Caption = "HP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDMP
                    lblItemDescMod.Caption = "MP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSP
                    lblItemDescMod.Caption = "SP +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSTR
                    lblItemDescMod.Caption = "Strength +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDDEF
                    lblItemDescMod.Caption = "Defence +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDMAGI
                    lblItemDescMod.Caption = "Magic +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDSPEED
                    lblItemDescMod.Caption = "Speed +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDDEX
                    lblItemDescMod.Caption = "Dexterity +" & Item(ItemNum).Data2
                    
                Case CHARM_TYPE_ADDCRIT
                    lblItemDescMod.Caption = "Critical +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDDROP
                    lblItemDescMod.Caption = "Drop +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDBLOCK
                    lblItemDescMod.Caption = "Block +" & Item(ItemNum).Data2 & "%"
                    
                Case CHARM_TYPE_ADDACCU
                    lblItemDescMod.Caption = "Accuracy +" & Item(ItemNum).Data2 & "%"
            End Select
        
        Case ITEM_TYPE_SKILL
            picItemDesc.Height = 65
            lblItemDescDur.Caption = ""
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 32
            Label10.Caption = "Skill Held"
            lblItemDescSpell.top = 48
            lblItemDescSpell.Caption = Trim(Skill(Item(ItemNum).Data1).Name)
            
        Case ITEM_TYPE_ARROW
            picItemDesc.Height = 49
            Label9.top = 64
            lblItemDescMod.top = 80
            Label10.top = 96
            lblItemDescSpell.top = 112
            lblItemDescDur.Caption = ItemDur & " / " & Item(ItemNum).Data1
            lblItemDescReq.top = 32
            lblItemDescReq.Caption = "Range: " & Item(ItemNum).Data2
    End Select
    
    'picItemDesc.Left = (x + 10)
    If Index > 14 Then
        picItemDesc.top = (Int(Index / 3) * 32)
    Else
        picItemDesc.top = ((Int(Index / 3) * 32) + 120)
    End If
    picItemDesc.Visible = True
End Sub

Private Sub lblItemDescSpell_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picItemDesc.Visible = False
End Sub

Private Sub lblItemDescMod_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picItemDesc.Visible = False
End Sub

Private Sub lblItemDescReq_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picItemDesc.Visible = False
End Sub

Private Sub lblItemDescDur_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picItemDesc.Visible = False
End Sub

Private Sub picPlayerInventory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerInventory.Drag vbBeginDrag
End Sub

Private Sub picPlayerInventory_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerInventory.Left - MouseXOffset, y + picPlayerInventory.top - MouseYOffset
    picScreen.SetFocus
End Sub

' *****************
' * Player Spells *
' *****************

Private Sub picSpell_DblClick(Index As Integer)
    If Player(MyIndex).Spells(Index + 1).Num > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & Index + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
                picScreen.SetFocus
            Else
                MsgMessage = "Still Moving"
                MessageColor = BrightRed
                MessageTime = GetTickCount
                iv = 0
                picScreen.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgMessage = "No Spell Detected"
        MessageColor = BrightRed
        MessageTime = GetTickCount
        iv = 0
        picScreen.SetFocus
        Exit Sub
    End If
End Sub

Private Sub picSpell_Click(Index As Integer)
    picScreen.SetFocus
End Sub

Private Sub picSpell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim SpellSlot As Long

    SpellSlot = Index + 1
    
    If GetPlayerSpell(MyIndex, SpellSlot) <= 0 Then
        picSpellDesc.Visible = False
        Exit Sub
    End If
    
    If SpellSlot <= 5 Then
        picSpellDesc.Left = 240
    Else
        picSpellDesc.Left = 75
    End If
    
    lblSpellDescName.Caption = Trim(Spell(Player(MyIndex).Spells(SpellSlot).Num).Name)
    lblSpellDescLevel.Caption = "Level: " & STR(Player(MyIndex).Spells(SpellSlot).Level)
    lblSpellDescExp.Caption = "Exp: " & STR(Player(MyIndex).Spells(SpellSlot).EXP)
    Select Case Spell(Player(MyIndex).Spells(SpellSlot).Num).Type
        Case SPELL_TYPE_STAT
            If Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_ADDHP Then
                lblSpellDescMod = "+" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " HP"
                lblSpellDescPer = "per Level"
            ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_ADDMP Then
                lblSpellDescMod = "+" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " MP"
                lblSpellDescPer = "per Level"
            ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_ADDSP Then
                lblSpellDescMod = "+" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " SP"
                lblSpellDescPer = "per Level"
            ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_SUBHP Then
                lblSpellDescMod = "-" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " HP"
                lblSpellDescPer = "per Level"
            ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_SUBMP Then
                lblSpellDescMod = "-" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " MP"
                lblSpellDescPer = "per Level"
            ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_STAT_SUBSP Then
                lblSpellDescMod = "-" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & " SP"
                lblSpellDescPer = "per Level"
            Else
                lblSpellDescMod = ""
                lblSpellDescPer = ""
            End If
            
        Case SPELL_TYPE_GIVEITEM
            lblSpellDescMod = "Gain: " & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2)
            lblSpellDescPer = Trim(Item(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1).Name)
            'If Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_CHANCE_CRIT Then
                'lblSpellDescMod = "+" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & "% Critical"
                'lblSpellDescPer = "per Level"
            'ElseIf Spell(Player(MyIndex).Spells(SpellSlot).Num).Data1 = SPELL_CHANCE_DROP Then
                'lblSpellDescMod = "+" & STR(Spell(Player(MyIndex).Spells(SpellSlot).Num).Data2) & "% Drop Chance"
                'lblSpellDescPer = "per Level"
            'Else
                'lblSpellDescMod = ""
                'lblSpellDescPer = ""
            'End If
    End Select
    picSpellDesc.Visible = True
End Sub

Private Sub chkPlayerSpellsPin_Click()
    If chkPlayerSpellsPin.Value = Checked Then
        tmrPlayerSpells.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerSpells.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub lblPlayerSpellsCancel_Click()
    picPlayerSpells.Visible = False
    tmrPlayerSpells.Enabled = False
    chkPlayerSpellsPin.Value = Unchecked
End Sub
    
Private Sub tmrPlayerSpells_Timer()
    picPlayerSpells.Visible = False
    tmrPlayerSpells.Enabled = False
End Sub

Private Sub picPlayerSpells_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerSpells.Drag vbBeginDrag
End Sub

Private Sub picPlayerSpells_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerSpells.Left - MouseXOffset, y + picPlayerSpells.top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub picPlayerSpellsBack_Click()
    picScreen.SetFocus
End Sub

Private Sub lstPlayerSpells_Click()
    If picScreen.Visible Then picScreen.SetFocus
End Sub

' *****************
' * Player Skills *
' *****************

Private Sub picSkill_DblClick(Index As Integer)
    If Player(MyIndex).Skills(Index + 1).Num > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                'Call SendData("cast" & SEP_CHAR & Index + 1 & SEP_CHAR & END_CHAR)
                'Player(MyIndex).Attacking = 1
                'Player(MyIndex).AttackTimer = GetTickCount
                'Player(MyIndex).CastedSpell = YES
                picScreen.SetFocus
            Else
                MsgMessage = "Still Moving"
                MessageColor = BrightRed
                MessageTime = GetTickCount
                iv = 0
                picScreen.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgMessage = "No Skill Detected"
        MessageColor = BrightRed
        MessageTime = GetTickCount
        iv = 0
        picScreen.SetFocus
        Exit Sub
    End If
End Sub

Private Sub picSkill_Click(Index As Integer)
    picScreen.SetFocus
End Sub

Private Sub picSkill_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim SkillSlot As Long

    SkillSlot = Index + 1
    
    If GetPlayerSkill(MyIndex, SkillSlot) <= 0 Then
        picSkillDesc.Visible = False
        Exit Sub
    End If
    
    If SkillSlot <= 5 Then
        picSkillDesc.Left = 240
    Else
        picSkillDesc.Left = 75
    End If
    
    lblSkillDescName.Caption = Trim(Skill(Player(MyIndex).Skills(SkillSlot).Num).Name)
    lblSkillDescLevel.Caption = "Level: " & STR(Player(MyIndex).Skills(SkillSlot).Level)
    lblSkillDescExp.Caption = "Exp: " & STR(Player(MyIndex).Skills(SkillSlot).EXP)
    Select Case Skill(Player(MyIndex).Skills(SkillSlot).Num).Type
        Case SKILL_TYPE_ATTRIBUTE
            If Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_ATTRIBUTE_STR Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & " Strength"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_ATTRIBUTE_DEF Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & " Defence"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_ATTRIBUTE_MAGI Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & " Magic"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_ATTRIBUTE_SPEED Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & " Speed"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_ATTRIBUTE_DEX Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & " Dexterity"
                lblSkillDescPer = "per Level"
            Else
                lblSkillDescMod = ""
                lblSkillDescPer = ""
            End If
            
        Case SKILL_TYPE_CHANCE
            If Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_CHANCE_CRIT Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & "% Critical"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_CHANCE_DROP Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & "% Drop Chance"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_CHANCE_BLOCK Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & "% Block Chance"
                lblSkillDescPer = "per Level"
            ElseIf Skill(Player(MyIndex).Skills(SkillSlot).Num).Data1 = SKILL_CHANCE_ACCU Then
                lblSkillDescMod = "+" & STR(Skill(Player(MyIndex).Skills(SkillSlot).Num).Data2) & "% Accuracy"
                lblSkillDescPer = "per Level"
            Else
                lblSkillDescMod = ""
                lblSkillDescPer = ""
            End If
    End Select
    picSkillDesc.Visible = True
End Sub

Private Sub chkPlayerSkillsPin_Click()
    If chkPlayerSkillsPin.Value = Checked Then
        tmrPlayerSkills.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerSkills.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub lblPlayerSkillsCancel_Click()
    picPlayerSkills.Visible = False
    tmrPlayerSkills.Enabled = False
    chkPlayerSkillsPin.Value = Unchecked
End Sub
    
Private Sub tmrPlayerSkills_Timer()
    picPlayerSkills.Visible = False
    tmrPlayerSkills.Enabled = False
End Sub

Private Sub picPlayerSkills_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerSkills.Drag vbBeginDrag
End Sub

Private Sub picPlayerSkills_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerSkills.Left - MouseXOffset, y + picPlayerSkills.top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub picPlayerSkillsBack_Click()
    picScreen.SetFocus
End Sub

' *****************
' * Player Quests *
' *****************

Private Sub tmrPlayerQuests_Timer()
    picPlayerQuests.Visible = False
    tmrPlayerQuests.Enabled = False
End Sub

Private Sub lblPlayerQuestQuit_Click()
    picPlayerQuests.Visible = False
    tmrPlayerQuests.Enabled = False
    chkPlayerQuestPin.Value = Unchecked
End Sub

Private Sub chkPlayerQuestPin_Click()
    If chkPlayerQuestPin.Value = Checked Then
        tmrPlayerQuests.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerQuests.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub picPlayerQuests_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerQuests.Drag vbBeginDrag
End Sub

Private Sub picPlayerQuests_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerQuests.Left - MouseXOffset, y + picPlayerQuests.top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub scrlPlayerQuestNum_Change()
Dim QuestSlot As Long
Dim QuestNum As Long

    QuestSlot = scrlPlayerQuestNum.Value
    QuestNum = Player(MyIndex).Quests(QuestSlot).Num

    lblPlayerQuestNum.Caption = STR(scrlPlayerQuestNum.Value)
    If QuestNum > 0 Then
        lblPlayerQuestName.Caption = Trim(Quest(QuestNum).Name)
        txtPlayerQuestDesc.Text = Trim(Quest(QuestNum).Description)
        If Quest(QuestNum).RewardValue > 1 Then
            lblPlayerQuestReward.Caption = STR(Quest(QuestNum).RewardValue) & "  x  " & Trim(Item(Quest(QuestNum).Reward).Name)
        Else
            lblPlayerQuestReward.Caption = Trim(Item(Quest(QuestNum).Reward).Name)
        End If
        lblPlayerQuestCount.Caption = STR(Player(MyIndex).Quests(QuestSlot).Count) & " / " & STR(Player(MyIndex).Quests(QuestSlot).Amount)
    Else
        lblPlayerQuestName.Caption = "No Quest"
        txtPlayerQuestDesc.Text = ""
        lblPlayerQuestReward.Caption = ""
        lblPlayerQuestCount.Caption = ""
    End If
    
    If picScreen.Visible Then picScreen.SetFocus
End Sub

Private Sub txtPlayerQuestDesc_Click()
    picScreen.SetFocus
End Sub

' :::::::::::::::::
' :: Player Chat ::
' :::::::::::::::::

Private Sub tmrPlayerChat_Timer()
    picPlayerChat.Visible = False
    tmrPlayerChat.Enabled = False
End Sub

Private Sub lblPlayerChatCancel_Click()
    picPlayerChat.Visible = False
    tmrPlayerChat.Enabled = False
    chkPlayerChatPin.Value = Unchecked
End Sub

Private Sub chkPlayerChatPin_Click()
    If chkPlayerChatPin.Value = Checked Then
        tmrPlayerChat.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrPlayerChat.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub picPlayerChat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picPlayerChat.Drag vbBeginDrag
End Sub

Private Sub picPlayerChat_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picPlayerChat.Left - MouseXOffset, y + picPlayerChat.top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub txtChat_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move (x / 15) + picPlayerChat.Left + txtChat.Left - MouseXOffset, (y / 15) + picPlayerChat.top + txtChat.top - MouseYOffset
    picScreen.SetFocus
End Sub

' ::::::::::::::::
' :: Npc Quests ::
' ::::::::::::::::

Private Sub cmdCancelQuest_Click()
    picNpcQuests.Visible = False
    lstNpcQuests.Clear
    cmdCompleteQuest.Visible = False
    cmdAbandonQuest.Visible = False
    QuestNpcNum = 0
    chkPlayerQuestPin.Value = Unchecked
    picScreen.SetFocus
End Sub

Private Sub lstNpcQuests_Click()
Dim QuestNumber As Long
Dim n As Long

    If QuestNpcNum = 0 Then Exit Sub

    QuestNumber = Npc(QuestNpcNum).QuestNPC(lstNpcQuests.ListIndex + 1)
    
    If QuestNumber > 0 Then
        If Quest(QuestNumber).ClassReq > 0 Then
            lblNpcQuestClass.Caption = Trim(Class(Quest(QuestNumber).ClassReq - 1).Name)
        Else
            lblNpcQuestClass.Caption = "All Classes"
        End If
        lblNpcQuestMin.Caption = STR(Quest(QuestNumber).LevelMin)
        lblNpcQuestMax.Caption = STR(Quest(QuestNumber).LevelMax)
        If Quest(QuestNumber).RewardValue > 1 Then
            lblNpcQuestReward.Caption = STR(Quest(QuestNumber).RewardValue) & "  x  " & Trim(Item(Quest(QuestNumber).Reward).Name)
        Else
            lblNpcQuestReward.Caption = Trim(Item(Quest(QuestNumber).Reward).Name)
        End If
        txtNpcQuestDesc.Text = Trim(Quest(QuestNumber).Description)
        
        ' Check for completed quests
        For n = 1 To MAX_PLAYER_QUESTS
            If (Player(MyIndex).Quests(n).Num = QuestNumber) And (Player(MyIndex).Quests(n).SetMap = GetPlayerMap(MyIndex)) And (Player(MyIndex).Quests(n).SetBy = QuestNpcNum) Then
                If Player(MyIndex).Quests(n).Amount <> Player(MyIndex).Quests(n).Count Then
                    cmdCompleteQuest.Visible = False
                    cmdAbandonQuest.Visible = True
                    picScreen.SetFocus
                    Exit Sub
                Else
                    cmdCompleteQuest.Visible = True
                    cmdAbandonQuest.Visible = False
                    picScreen.SetFocus
                    Exit Sub
                End If
            End If
            cmdCompleteQuest.Visible = False
            cmdAbandonQuest.Visible = False
        Next n
    Else
        lblNpcQuestClass.Caption = ""
        lblNpcQuestMin.Caption = ""
        lblNpcQuestMax.Caption = ""
        lblNpcQuestReward.Caption = ""
        txtNpcQuestDesc.Text = ""
        cmdAbandonQuest.Visible = False
        cmdCompleteQuest.Visible = False
    End If
    picScreen.SetFocus
End Sub

Private Sub cmdAcceptQuest_Click()
    Call SendData("acceptquest" & SEP_CHAR & QuestNpcNum & SEP_CHAR & (lstNpcQuests.ListIndex + 1) & SEP_CHAR & END_CHAR)
    picScreen.SetFocus
End Sub

Private Sub cmdCompleteQuest_Click()
    Call SendData("completequest" & SEP_CHAR & QuestNpcNum & SEP_CHAR & (lstNpcQuests.ListIndex + 1) & SEP_CHAR & END_CHAR)
    picScreen.SetFocus
End Sub

Private Sub cmdAbandonQuest_Click()
    Call SendData("abandonquest" & SEP_CHAR & QuestNpcNum & SEP_CHAR & (lstNpcQuests.ListIndex + 1) & SEP_CHAR & END_CHAR)
    picScreen.SetFocus
End Sub

Private Sub picNpcQuests_Click()
    picScreen.SetFocus
End Sub

' ****************
' * Shop Trading *
' ****************

Private Sub cmdShopDeal_Click()
    If lstShopTrade.ListIndex >= 0 Then
        If Not lstShopTrade.Text = "" And Not lstShopTrade.Text = "Empty Trade Slot" Then
            Call SendData("traderequest" & SEP_CHAR & Player(MyIndex).ShopNum & SEP_CHAR & lstShopTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
            picScreen.SetFocus
        Else
            MsgMessage = "No Trade Available"
            MessageColor = BrightRed
            MessageTime = GetTickCount
            iv = 0
            picScreen.SetFocus
        End If
    End If
End Sub

Private Sub cmdShopFixItems_Click()
Dim i As Long

    frmCClient.lstFixItem.Clear
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmCClient.lstFixItem.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmCClient.lstFixItem.AddItem "Unused Slot"
        End If
    Next i
    lstFixItem.ListIndex = 0
    picFixItems.Visible = True
    picScreen.SetFocus
End Sub

Private Sub cmdItemFix_Click()
    Call SendData("fixitem" & SEP_CHAR & lstFixItem.ListIndex + 1 & SEP_CHAR & END_CHAR)
    picScreen.SetFocus
End Sub

Private Sub cmdShopCancel_Click()
    picFixItems.Visible = False
    lstFixItem.Clear
    picShopTrade.Visible = False
    Player(MyIndex).ShopNum = 0
    picScreen.SetFocus
End Sub

Private Sub picShopTrade_Click()
    picScreen.SetFocus
End Sub

Private Sub lstShopTrade_GotFocus()
    picScreen.SetFocus
End Sub

Private Sub lstFixItem_GotFocus()
    picScreen.SetFocus
End Sub
