VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13470
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   898
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   976
   StartUpPosition =   2  'CenterScreen
   Tag             =   " "
   Visible         =   0   'False
   Begin VB.PictureBox picButton 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   15
      Left            =   14040
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   9240
      Width           =   480
   End
   Begin VB.PictureBox picButton 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   11
      Left            =   14040
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picQuestAccept 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   1320
      ScaleHeight     =   2115
      ScaleWidth      =   7155
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Label lblQuestName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quest Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2880
         TabIndex        =   210
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label lblDecline 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decline"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   209
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lblAccept 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   120
         TabIndex        =   208
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lblQuestMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quest Start Message"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   207
         Top             =   480
         Width           =   6915
      End
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   30
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   834
      TabIndex        =   134
      Top             =   12360
      Visible         =   0   'False
      Width           =   12540
      Begin VB.CheckBox chkLayers 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Eye Dropper (Shift + LMouse)"
         Top             =   120
         Width           =   540
      End
      Begin VB.CheckBox chkTilesets 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Eye Dropper (Shift + LMouse)"
         Top             =   120
         Width           =   540
      End
      Begin VB.CheckBox chkDimLayers 
         BackColor       =   &H0080C0FF&
         Caption         =   "Dim Layers"
         Height          =   225
         Left            =   9780
         TabIndex        =   147
         ToolTipText     =   "Will dim tiles of layers that are below your current layer."
         Top             =   180
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0080C0FF&
         Height          =   420
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Delete all content of this map. "
         Top             =   540
         Width           =   420
      End
      Begin VB.CommandButton cmdRevert 
         BackColor       =   &H0080C0FF&
         Height          =   420
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Revert/Cancel all changes to this map."
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton cmdProperties 
         BackColor       =   &H0080C0FF&
         Height          =   420
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Map Properties"
         Top             =   540
         Width           =   420
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Height          =   420
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Save Map"
         Top             =   120
         Width           =   420
      End
      Begin VB.CheckBox chkDrawEvents 
         BackColor       =   &H0080C0FF&
         Caption         =   "Draw Events"
         Height          =   225
         Left            =   9780
         TabIndex        =   142
         ToolTipText     =   "Draw white square around events "
         Top             =   420
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkShowAttributes 
         BackColor       =   &H0080C0FF&
         Caption         =   "Show Attributes"
         Height          =   225
         Left            =   8100
         TabIndex        =   141
         ToolTipText     =   "Show attributes like Block, Warp etc. "
         Top             =   660
         Width           =   1740
      End
      Begin VB.CheckBox chkGrid 
         BackColor       =   &H0080C0FF&
         Caption         =   "Show Grid"
         Height          =   225
         Left            =   8100
         TabIndex        =   140
         Top             =   420
         Width           =   1395
      End
      Begin VB.CheckBox chkTilePreview 
         BackColor       =   &H0080C0FF&
         Caption         =   "Tile Preview"
         Height          =   225
         Left            =   8100
         TabIndex        =   139
         Top             =   180
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkEyeDropper 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Eye Dropper (Shift + LMouse)"
         Top             =   120
         Width           =   540
      End
      Begin VB.CheckBox mapPreviewSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Map Preview - Docked"
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(MWheel Scroll)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   180
         Left            =   4560
         TabIndex        =   155
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label lblLayers 
         BackStyle       =   0  'Transparent
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   154
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lblHotMapPreview 
         BackStyle       =   0  'Transparent
         Caption         =   "(Ctrl+M)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   180
         Left            =   960
         TabIndex        =   152
         Top             =   825
         Width           =   570
      End
      Begin VB.Label lblHotEye 
         BackStyle       =   0  'Transparent
         Caption         =   "(Shift+ LMouse)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   180
         Left            =   2760
         TabIndex        =   151
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label lblHotTilesets 
         BackStyle       =   0  'Transparent
         Caption         =   "(Middle MBtn)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   180
         Left            =   6360
         TabIndex        =   150
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblTilesets 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Tileset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   149
         Top             =   660
         Width           =   870
      End
      Begin VB.Label lblTilePreview 
         BackStyle       =   0  'Transparent
         Caption         =   "Eye Dropper"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   138
         Top             =   645
         Width           =   870
      End
      Begin VB.Label lblMapPreview 
         BackStyle       =   0  'Transparent
         Caption         =   "Map Preview"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   136
         Top             =   660
         Width           =   930
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   14160
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   225
      TabStop         =   0   'False
      Top             =   12240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   13860
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   12900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picButton 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   13
      Left            =   15780
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   13380
      Visible         =   0   'False
      Width           =   480
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   8
         Left            =   360
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   480
         Begin VB.PictureBox deleteThis 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   15000
            Index           =   11
            Left            =   300
            ScaleHeight     =   1000
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1000
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   15000
         End
      End
   End
   Begin VB.PictureBox picQuest 
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   8880
      ScaleHeight     =   4875
      ScaleWidth      =   3315
      TabIndex        =   217
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ListBox lstQuests 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   4050
         Left            =   120
         TabIndex        =   218
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblSelectOne 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select one to view details."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   220
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label lblQuestLog 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Log"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1140
         TabIndex        =   219
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picQuestDesc 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   4935
      Left            =   5340
      ScaleHeight     =   4875
      ScaleWidth      =   3435
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtQuestTask 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   213
         Text            =   "frmMain.frx":038A
         Top             =   2460
         Width           =   3135
      End
      Begin VB.CommandButton btnQuestCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quit This Quest"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   4500
         Width           =   3255
      End
      Begin VB.Timer tmrRUSure 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2640
         Top             =   120
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1035
         TabIndex        =   216
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblQuestDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quest Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   215
         Top             =   480
         Width           =   3135
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0FFFF&
         Index           =   1
         X1              =   180
         X2              =   3300
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblCurrentTask 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Task(s):"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   840
         TabIndex        =   214
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0FFFF&
         Index           =   0
         X1              =   60
         X2              =   3540
         Y1              =   4380
         Y2              =   4380
      End
   End
   Begin VB.PictureBox picGuild 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   11760
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   11040
      Visible         =   0   'False
      Width           =   2940
      Begin VB.PictureBox picGuild_No 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   60
         ScaleHeight     =   273
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lblYouAre 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You are not in a guild!"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   330
            TabIndex        =   205
            Top             =   2040
            Width           =   2205
         End
      End
      Begin VB.ListBox lstGuild 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2130
         Left            =   240
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   600
         Width           =   2460
      End
      Begin VB.Label lblChangeAccess 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Access"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   203
         Top             =   3000
         Width           =   2715
      End
      Begin VB.Label lblResign 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resign"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   202
         Top             =   3720
         Width           =   2715
      End
      Begin VB.Label lblInvite 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invite"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   201
         Top             =   3240
         Width           =   2715
      End
      Begin VB.Label lblRemove 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   200
         Top             =   3480
         Width           =   795
      End
      Begin VB.Label lblGuildName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   199
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   11880
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   122
      Top             =   11040
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   123
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   125
         Top             =   210
         Width           =   2805
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   124
         Top             =   1800
         Width           =   2640
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   11760
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   118
      Top             =   11640
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   119
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   121
         Top             =   210
         Width           =   2805
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   120
         Top             =   1800
         Width           =   2640
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picOptionSwearFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   15540
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   12420
      Width           =   735
   End
   Begin VB.PictureBox picOptionWeather 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   14685
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   12390
      Width           =   735
   End
   Begin VB.PictureBox picOptionAutoTile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   14685
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   12750
      Width           =   735
   End
   Begin VB.PictureBox picOptionDebug 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   14685
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   13470
      Width           =   735
   End
   Begin VB.PictureBox picOptionBlood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   14685
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   13110
      Width           =   735
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   15720
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   12780
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picSkills 
      Appearance      =   0  'Flat
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   8130
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label lblSkills 
         BackStyle       =   0  'Transparent
         Caption         =   "Skills"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   240
         Left            =   1740
         TabIndex        =   193
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   3
         Left            =   2640
         TabIndex        =   192
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   191
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   190
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   189
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   188
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   187
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   186
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   185
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   184
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   183
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   182
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   4
         Left            =   480
         TabIndex        =   181
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   180
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   179
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   5
         Left            =   2640
         TabIndex        =   178
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   177
         Top             =   3720
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   176
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   6
         Left            =   480
         TabIndex        =   175
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   7
         Left            =   2640
         TabIndex        =   174
         Top             =   3720
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   7
         Left            =   2640
         TabIndex        =   173
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   7
         Left            =   2640
         TabIndex        =   172
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   171
         Top             =   4560
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   170
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   8
         Left            =   480
         TabIndex        =   169
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   9
         Left            =   2640
         TabIndex        =   168
         Top             =   4560
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   9
         Left            =   2640
         TabIndex        =   167
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   9
         Left            =   2640
         TabIndex        =   166
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   165
         Top             =   5400
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   164
         Top             =   5160
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   10
         Left            =   480
         TabIndex        =   163
         Top             =   4920
         Width           =   450
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   11
         Left            =   2640
         TabIndex        =   162
         Top             =   5400
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   11
         Left            =   2640
         TabIndex        =   161
         Top             =   5160
         Width           =   465
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   11
         Left            =   2640
         TabIndex        =   160
         Top             =   4920
         Width           =   450
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   159
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   1
         Left            =   2640
         TabIndex        =   158
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   157
         Top             =   1560
         Width           =   450
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   0  'User
      ScaleWidth      =   484
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7740
      Visible         =   0   'False
      Width           =   7260
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   480
         Width           =   7155
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   7260
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   7740
      Visible         =   0   'False
      Width           =   7260
      Begin VB.TextBox txtDialogue 
         Height          =   315
         Left            =   2520
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   62
         Top             =   1665
         Width           =   285
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player has requested a trade. Would you like to accept?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   59
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3405
         TabIndex        =   58
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3360
         TabIndex        =   61
         Top             =   1440
         Width           =   345
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   2400
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3840
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   4500
         Width           =   1815
      End
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   3960
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   360
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1530
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2700
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image ImgFix 
         Height          =   315
         Left            =   1890
         Top             =   3840
         Width           =   375
      End
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10020
      Left            =   120
      ScaleHeight     =   668
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   864
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12960
      Begin VB.PictureBox picSpells 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9300
         ScaleHeight     =   270
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   9
         Left            =   11160
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   10
         Left            =   10560
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   9480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   16
         Left            =   10560
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   12
         Left            =   9960
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   9480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   14
         Left            =   11760
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   9480
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   8760
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   3
         Left            =   9960
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   4
         Left            =   11760
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   5
         Left            =   8760
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   9480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   6
         Left            =   11160
         ScaleHeight     =   29
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   9480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   7
         Left            =   9360
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   9480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   9360
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   8940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picHotbar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5040
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   476
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   7140
      End
      Begin VB.PictureBox picGUI_Vitals_Base 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   120
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   254
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   3810
         Begin VB.Label lblHP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   19
            Top             =   135
            Width           =   1845
         End
         Begin VB.Label lblMP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   18
            Top             =   465
            Width           =   1845
         End
         Begin VB.Label lblEXP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   17
            Top             =   795
            Width           =   1845
         End
         Begin VB.Image imgHPBar 
            Height          =   240
            Left            =   105
            Top             =   135
            Width           =   3615
         End
         Begin VB.Image imgMPBar 
            Height          =   240
            Left            =   105
            Top             =   465
            Width           =   3615
         End
         Begin VB.Image imgEXPBar 
            Height          =   240
            Left            =   120
            Top             =   795
            Width           =   3615
         End
      End
      Begin VB.PictureBox picFriends 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9300
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
         Begin VB.ListBox lstFriends 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2550
            Left            =   300
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   600
            Width           =   2340
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Friends"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   60
            TabIndex        =   194
            Top             =   180
            Width           =   2835
         End
         Begin VB.Label lblRemoveFriend 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Friend"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   60
            TabIndex        =   55
            Top             =   3600
            Width           =   2805
         End
         Begin VB.Label lblAddFriend 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add Friend"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   900
            TabIndex        =   54
            Top             =   3300
            Width           =   1125
         End
      End
      Begin VB.PictureBox picFoes 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9300
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
         Begin VB.ListBox lstFoes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2550
            Left            =   300
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   600
            Width           =   2340
         End
         Begin VB.Label lblFoes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Foes"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   75
            TabIndex        =   196
            Top             =   180
            Width           =   2805
         End
         Begin VB.Label lblAddFoe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add Foe"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1140
            TabIndex        =   80
            Top             =   3300
            Width           =   795
         End
         Begin VB.Label lblRemoveFoe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Foe"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   960
            TabIndex        =   79
            Top             =   3600
            Width           =   1155
         End
      End
      Begin VB.PictureBox picInventory 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9300
         ScaleHeight     =   269.004
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox picCharacter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4200
         Left            =   9300
         ScaleHeight     =   280
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3960
         Visible         =   0   'False
         Width           =   2925
         Begin VB.PictureBox picFace 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   735
            ScaleHeight     =   96
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   570
            Width           =   1440
         End
         Begin VB.Label lblCharLevel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lv: 1"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1200
            TabIndex        =   74
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label lblCharName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empty"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1170
            TabIndex        =   44
            Top             =   150
            Width           =   660
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   720
            TabIndex        =   43
            Top             =   2880
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   2040
            TabIndex        =   42
            Top             =   2880
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   720
            TabIndex        =   41
            Top             =   3120
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   2040
            TabIndex        =   40
            Top             =   3120
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   720
            TabIndex        =   39
            Top             =   3360
            Width           =   315
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   38
            Top             =   2880
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   2520
            TabIndex        =   37
            Top             =   2880
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   1200
            TabIndex        =   36
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   2520
            TabIndex        =   35
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   1200
            TabIndex        =   34
            Top             =   3360
            Width           =   120
         End
         Begin VB.Label lblPoints 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   2280
            TabIndex        =   33
            Top             =   3360
            Width           =   315
         End
         Begin VB.Label lblStr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   32
            Top             =   2880
            Width           =   360
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   31
            Top             =   3120
            Width           =   435
         End
         Begin VB.Label lblInt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Int:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   3360
            Width           =   360
         End
         Begin VB.Label lblSpi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spi:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   29
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label lblAgi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agi:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   28
            Top             =   2880
            Width           =   390
         End
         Begin VB.Label lblPoint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   27
            Top             =   3360
            Width           =   675
         End
      End
      Begin VB.PictureBox picParty 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   9300
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   191
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4410
         Visible         =   0   'False
         Width           =   2865
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   4
            Left            =   105
            Top             =   2760
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   4
            Left            =   105
            Top             =   2625
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   3
            Left            =   105
            Top             =   2025
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   3
            Left            =   105
            Top             =   1890
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   2
            Left            =   105
            Top             =   1320
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   2
            Left            =   105
            Top             =   1170
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   555
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   1
            Left            =   105
            Top             =   420
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Label lblPartyLeave 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1575
            TabIndex        =   51
            Top             =   3165
            Width           =   1095
         End
         Begin VB.Label lblPartyInvite 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   375
            TabIndex        =   50
            Top             =   3165
            Width           =   1095
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   255
            TabIndex        =   49
            Top             =   2355
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   255
            TabIndex        =   48
            Top             =   1620
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   47
            Top             =   885
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   255
            TabIndex        =   46
            Top             =   150
            Width           =   2415
         End
      End
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   9300
         ScaleHeight     =   160
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   5760
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   9300
         ScaleHeight     =   270
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
         Begin VB.PictureBox picOptionMusic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.PictureBox picOptionSound 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.PictureBox picOptionLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   960
            Width           =   735
         End
         Begin VB.PictureBox picOptionGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1320
            Width           =   735
         End
         Begin VB.PictureBox picOptionTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.PictureBox picOptionWASD 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   2040
            Width           =   735
         End
         Begin VB.PictureBox picOptionMouse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   2400
            Width           =   735
         End
         Begin VB.PictureBox picOptionBattleMusic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   2760
            Width           =   735
         End
         Begin VB.PictureBox picOptionNpcVitals 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   3480
            Width           =   735
         End
         Begin VB.PictureBox picOptionPlayerVitals 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label lblMusic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Music"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   116
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblSound 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sound"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   115
            Top             =   600
            Width           =   600
         End
         Begin VB.Label lblGuilds 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guilds"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   114
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label lblLevels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Levels"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   113
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblFKeys 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F Keys"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   112
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblMouse 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   111
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label lblTitles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Titles"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   110
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label lblNMobVitals 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mob Vitals"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   109
            Top             =   3480
            Width           =   1035
         End
         Begin VB.Label lblBattleMusic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battle Music"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   108
            Top             =   2760
            Width           =   1200
         End
         Begin VB.Label lblPlayerVitals 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Player Vitals"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   107
            Top             =   3120
            Width           =   1290
         End
      End
      Begin VB.PictureBox picTitles 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9300
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4110
         Visible         =   0   'False
         Width           =   2925
         Begin VB.ListBox lstTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2550
            Left            =   300
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   600
            Width           =   2340
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Titles"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   60
            TabIndex        =   195
            Top             =   180
            Width           =   2805
         End
         Begin VB.Label lblDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "None."
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   240
            TabIndex        =   84
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   855
            TabIndex        =   83
            Top             =   3240
            Width           =   1215
         End
      End
      Begin VB.PictureBox picEventChat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2790
         Left            =   120
         ScaleHeight     =   186
         ScaleMode       =   0  'User
         ScaleWidth      =   482
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   7125
         Visible         =   0   'False
         Width           =   7230
         Begin VB.PictureBox picChatFace 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1500
            Left            =   120
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   100
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lblEventChat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "[Text]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1515
            Left            =   1680
            TabIndex        =   128
            Top             =   120
            Width           =   5535
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 1]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   133
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 2]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   132
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 3]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   3
            Left            =   3240
            TabIndex        =   131
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 4]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   4
            Left            =   3240
            TabIndex        =   130
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblEventChatContinue 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Continue..."
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Left            =   6000
            TabIndex        =   129
            Top             =   2400
            Width           =   1095
         End
      End
      Begin VB.PictureBox picChatbox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   484
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   7740
         Width           =   7260
         Begin VB.TextBox txtMyChat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   600
            MaxLength       =   512
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1890
            Width           =   6585
         End
         Begin RichTextLib.RichTextBox txtChat 
            Height          =   1755
            Left            =   60
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   60
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   3096
            _Version        =   393217
            BackColor       =   -2147483647
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmMain.frx":03A8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox picScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5760
         Left            =   0
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   0
         Width           =   7680
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.Label lblSwearFilter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Swear Filter"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   13380
      TabIndex        =   95
      Top             =   12360
      Width           =   1200
   End
   Begin VB.Label lblWeather 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weather"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12885
      TabIndex        =   94
      Top             =   12390
      Width           =   855
   End
   Begin VB.Label lblAutoTile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Tile"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12885
      TabIndex        =   93
      Top             =   12750
      Width           =   915
   End
   Begin VB.Label lblDebug 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12885
      TabIndex        =   92
      Top             =   13470
      Width           =   600
   End
   Begin VB.Label lblBlood 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blood"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12885
      TabIndex        =   91
      Top             =   13110
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private LastX As Long
Private LastY As Long
Private taskBarClick As Boolean

Private WithEvents cSubclasserHooker As cSelfSubHookCallback
Attribute cSubclasserHooker.VB_VarHelpID = -1

Private Sub btnCloseQuestLog_Click()
    picQuest.Visible = False
    picQuestDesc.Visible = False
    CurButton_Main = 0
    LastButton_Main = 0
    Call ResetMainButtons
End Sub

Private Sub btnQuestCancel_Click()
Dim Index As Long
    If btnQuestCancel.Caption = "Quit This Quest" Then
        btnQuestCancel.Caption = "Are You Sure You Want To Quit? (3)"
        tmrRUSure.Enabled = True
    Else
        'send a request to quit the selected quest
        btnQuestCancel.Caption = "Quit This Quest"
        Index = FindQuest(lstQuests.List(lstQuests.ListIndex))
        If Index < 1 Then Exit Sub
        picQuestDesc.Visible = False
        Call SendRequestQuitQuest(Index)
    End If
End Sub

Private Sub chkDimLayers_Click()
    redrawMapCache = True
End Sub

Private Sub chkEyeDropper_Click()
    If frmMain.chkEyeDropper.Value Then
        frmMain.chkEyeDropper.Picture = LoadResPicture("EYE_DOWN", vbResBitmap)
    Else
        frmMain.chkEyeDropper.Picture = LoadResPicture("EYE_UP", vbResBitmap)
    End If
End Sub

Public Sub chkLayers_Click()
    If frmMain.chkLayers.Value Then
        frmMain.chkLayers.Picture = LoadResPicture("LAYERS_DOWN", vbResBitmap)
        layersActive = True
    Else
        frmMain.chkLayers.Picture = LoadResPicture("LAYERS_UP", vbResBitmap)
        layersActive = False
    End If
End Sub

Private Sub chkTilePreview_Click()
    CurX = 0
    CurY = 0
End Sub

Private Sub chkTilesets_Click()
    If frmMain.chkTilesets.Value Then
        frmMain.chkTilesets.Picture = LoadResPicture("TILESETS_DOWN", vbResBitmap)
        displayTilesets = True
    Else
        frmMain.chkTilesets.Picture = LoadResPicture("TILESETS_UP", vbResBitmap)
        displayTilesets = False
    End If
End Sub

Private Sub cmdDelete_Click()
    If AlertMsg("Are you sure you want to erase this map?", False, False) = YES Then
        Call ClearMap
        Call MapEditorSave
        redrawMapCache = True
    End If
End Sub

Private Sub cmdProperties_Click()
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Load the values
    MapPropertiesInit
    
    ' Update the 1stnpcs list Index so it is selected
    frmEditor_MapProperties.lstNpcs.ListIndex = 0
    
    ' Show the form
    frmEditor_MapProperties.Show
    
    ' Lock map editor open til map properties is closed
    cmdSave.Enabled = False
    cmdRevert.Enabled = False
    Exit Sub
' Error handler
ErrorHandler:
    HandleError "cmdProperties_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdRevert_Click()
    LeaveMapEditorMode True
End Sub

Private Sub cmdSave_Click()
    Call MapEditorSave
    LeaveMapEditorMode True
End Sub

Public Sub SubDaFocus(hWnd As Long)
    If cSubclasserHooker.ssc_Subclass(hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg hWnd, eMsgWhen.MSG_BEFORE, WM_ACTIVATEAPP, WM_NCACTIVATE, WM_MOVE
    End If
End Sub

Public Sub UnsubDaFocus(hWnd As Long)
    cSubclasserHooker.ssc_UnSubclass hWnd
End Sub

Private Sub Form_Load()
    If cSubclasserHooker Is Nothing Then
        Set cSubclasserHooker = New cSelfSubHookCallback
    End If

    If cSubclasserHooker.ssc_Subclass(Me.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.hWnd, eMsgWhen.MSG_BEFORE, WM_ACTIVATEAPP, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_CAPTURECHANGED, WM_GETMINMAXINFO, WM_MOUSEWHEEL, WM_NCACTIVATE, WM_MOVE
    End If
    
    If cSubclasserHooker.ssc_Subclass(Me.picMapEditor.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.picMapEditor.hWnd, eMsgWhen.MSG_BEFORE, WM_ACTIVATEAPP, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_CAPTURECHANGED, WM_GETMINMAXINFO
    End If
    If cSubclasserHooker.ssc_Subclass(Me.mapPreviewSwitch.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.mapPreviewSwitch.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    If cSubclasserHooker.ssc_Subclass(Me.chkEyeDropper.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.chkEyeDropper.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    If cSubclasserHooker.ssc_Subclass(Me.cmdSave.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.cmdSave.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    If cSubclasserHooker.ssc_Subclass(Me.cmdRevert.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.cmdRevert.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    If cSubclasserHooker.ssc_Subclass(Me.cmdDelete.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.cmdDelete.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    If cSubclasserHooker.ssc_Subclass(Me.cmdProperties.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.cmdProperties.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
    
    SetIcon
End Sub

Private Sub Form_Paint()
    If FormVisible("frmCharEditor") Then
        frmCharEditor.Show
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If MyIndex > 0 Then
        If Moral(Map.Moral).CanPK = 1 Or Moral(Map.Moral).DropItems = 1 Or Moral(Map.Moral).LoseExp = 1 Or GetPlayerPK(MyIndex) = YES Then
            If AlertMsg("Are you sure you want to logout? You will remain logged in, please find a safe spot to logout!", False, False) = NO Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    
    If Not readyToExit Then
        Cancel = True
        Me.Visible = False
    Else
        cSubclasserHooker.ssc_UnSubclass Me.picMapEditor.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.mapPreviewSwitch.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.chkEyeDropper.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.cmdSave.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.cmdRevert.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.cmdDelete.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.cmdProperties.hWnd
        cSubclasserHooker.ssc_UnSubclass Me.hWnd
        Set cSubclasserHooker = Nothing
    End If
    
    If InGame Then
        IsLogging = True
        LogoutGame
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_QueryUnload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    
    ' Reset all buttons
    Call ResetMainButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picCurrency.Visible = False
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    AcceptTrade
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgFix_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InShop = 0 Then Exit Sub
    If Shop(InShop).CanFix = 0 Then Exit Sub
    
    TryingToFixItem = True
    
    AddText "Double-click on the item in your inventory you wish to fix.", BrightGreen
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgFix_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "Click on the item in the shop you wish to buy.", White
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    
    AddText "Double-click on the item in your inventory you wish to sell.", BrightGreen
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblAccept_Click()
Dim buffer As clsBuffer
    If Not QuestRequest > 0 Then
        picQuestAccept.Visible = False
        Exit Sub
    End If
    
    Set buffer = New clsBuffer
        buffer.WriteLong CAcceptQuest
        buffer.WriteLong QuestRequest
        Call SendData(buffer.ToArray())
        QuestRequest = 0
        picQuestAccept.Visible = False
    Set buffer = Nothing
End Sub

Private Sub lblAddFriend_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Add Friend", "Who do you want to add as a friend?", DIALOGUE_TYPE_ADDFRIEND, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblAddFriend_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblAddFoe_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Add Foe", "Who do you want to add as a foe?", DIALOGUE_TYPE_ADDFOE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblAddFoe_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChoices_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong Index
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    
    Call ClearChatButton(Index)
    InEvent = False
    Audio.PlaySound ButtonClick
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblChoices_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChoices_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(Index)
    If frmMain.lblChoices(Index).Visible = False Then Exit Sub
    If frmMain.lblChoices.Item(Index).ForeColor = vbYellow Then Exit Sub
    frmMain.lblChoices.Item(Index).ForeColor = vbYellow
    Audio.PlaySound ButtonHover
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblChoices_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ClearChatButton(Index As Integer)
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To 4
        If frmMain.lblChoices.Item(I).ForeColor = vbYellow And Not Index = I Then
            frmMain.lblChoices.Item(I).ForeColor = &H80000003
        End If
    Next
    
    frmMain.lblEventChatContinue.ForeColor = &H80000003
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChatButton", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ClearButtons()
    LastButton_Main = 0
    ResetOptionButtons
    Call ResetMainButtons
End Sub

Private Sub lblDecline_Click()
    QuestRequest = 0
    picQuestAccept.Visible = False
End Sub

Private Sub lblEventChatContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmMain.lblEventChatContinue.Visible = False Then Exit Sub
    If frmMain.lblEventChatContinue.ForeColor = vbYellow Then Exit Sub
    frmMain.lblEventChatContinue.ForeColor = vbYellow
    Audio.PlaySound ButtonHover
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblEventChatContinue_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub lblEventChatContinue_Click()
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong 0
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    InEvent = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblEventChatContinue_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEventChat()
    Dim I As Long
    
    If AnotherChat = 1 Then
        For I = 1 To 4
            frmMain.lblChoices(I).Visible = False
        Next
        
        frmMain.lblEventChat.Caption = ""
        frmMain.lblEventChatContinue.Visible = False
    ElseIf AnotherChat = 2 Then
        For I = 1 To 4
            frmMain.lblChoices(I).Visible = False
        Next
        
        frmMain.lblEventChat.Visible = False
        frmMain.lblEventChatContinue.Visible = False
        EventChatTimer = timeGetTime + 100
    Else
        frmMain.picEventChat.Visible = False
        frmMain.picChatbox.Visible = True
    End If
End Sub

Private Sub lblEquipCharName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblGuildRemove_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Guild Remove", "Who do you want to remove from the guild?", DIALOGUE_TYPE_GUILDREMOVE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblGuildRemove_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblRemoveFriend_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Remove Friend", "What friend do you want to remove?", DIALOGUE_TYPE_REMOVEFRIEND, True
    
    If (lstFriends.ListIndex + 1) > 0 And lstFriends.ListIndex + 1 <= MAX_PEOPLE Then
        txtDialogue.text = Trim$(Player(MyIndex).Friends(lstFriends.ListIndex + 1).Name)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblRemoveFriend_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblRemoveFoe_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Remove Foe", "What foe do you want to remove?", DIALOGUE_TYPE_REMOVEFOE, True
    
    If (lstFoes.ListIndex + 1) > 0 And lstFoes.ListIndex + 1 <= MAX_PEOPLE Then
        txtDialogue.text = Trim$(Player(MyIndex).Foes(lstFoes.ListIndex + 1).Name)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblRemoveFoe_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChangeAccess_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Change Guild Access", "What access would you like to change this user to?", DIALOGUE_TYPE_CHANGEGUILDACCESS, True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblChangeAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picCurrency.Visible = False
    txtCurrency.text = vbNullString
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    CloseTrade
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblLeaveBank_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    CloseBank
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblLeaveBank_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgLeaveShop_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InShop = 0 Then Exit Sub
    CloseShop
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsNumeric(txtCurrency.text) Then
        Select Case CurrencyMenu
            Case 1 ' Drop item
                SendDropItem TmpCurrencyItem, val(txtCurrency.text)
            Case 2 ' Deposit item
                DepositItem TmpCurrencyItem, val(txtCurrency.text)
            Case 3 ' withdraw item
                WithdrawItem TmpCurrencyItem, val(txtCurrency.text)
            Case 4 ' Offer trade item
                TradeItem TmpCurrencyItem, val(txtCurrency.text)
        End Select
    Else
        AddText "Please enter a valid amount.", BrightRed
        Exit Sub
    End If
    
    picCurrency.Visible = False
    TmpCurrencyItem = 0
    txtCurrency.text = vbNullString
    CurrencyMenu = 0 ' Clear
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblDialogue_Button_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Call the handler
    DialogueHandler Index
    
    txtDialogue.text = vbNullString
    picDialogue.Visible = False
    DialogueIndex = 0
    SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Party Invite", "Who do you want to invite to the party?", DIALOGUE_TYPE_PARTYINVITE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblPartyLeave_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Party.num > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblGuildInvite_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Dialogue "Guild Invite", "Who do you want to invite to the guild?", DIALOGUE_TYPE_GUILDINVITE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblGuildInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblResign_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    RequestGuildResign
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblResign_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
End Sub

Private Sub lblSpellName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
End Sub

Private Sub lblTrainStat_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstQuests_Click()
Dim Index As Long, CID As Long, TID As Long
Dim I As Long, II As Long, TmpStr As String
    If Not lstQuests.ListIndex > -1 Then
        picQuestDesc.Visible = False
        Exit Sub
    End If
    Index = FindQuest(lstQuests.List(lstQuests.ListIndex))
    If Index < 1 Then
        picQuestDesc.Visible = False
        Exit Sub
    End If
    
    CID = Player(MyIndex).QuestCLI(Index)
    TID = Player(MyIndex).QuestTask(Index)
    
    picQuestDesc.Visible = True
    lblQuestDesc.Caption = Trim$(Quest(Index).Description)
    txtQuestTask.text = vbNullString
    
    If CID < 1 Or TID < 1 Then
        If Not Player(MyIndex).QuestCompleted(Index) Then
            txtQuestTask.text = "Not Started..."
        Else
            txtQuestTask.text = "You have previously completed this quest."
        End If
        btnQuestCancel.Visible = False
        Exit Sub
    Else
        btnQuestCancel.Visible = True
    End If
    
    If TID - 1 > 0 Then
        For I = TID - 1 To 1 Step -1
            With Quest(Index).CLI(CID).Action(I)
                If Not .ActionID >= 1 Or Not .ActionID <= 4 Then Exit For
                II = II + 1
                If II > 1 Then TmpStr = vbNewLine Else TmpStr = vbNullString
                Select Case .ActionID
                    Case TASK_GATHER
                        txtQuestTask.text = txtQuestTask.text & TmpStr & "Gather " & .amount & " " & Trim$(Item(.MainData).Name)
                    Case TASK_KILL
                        txtQuestTask.text = txtQuestTask.text & TmpStr & "Kill " & .amount & " " & Trim$(NPC(.MainData).Name)
                    Case TASK_GETSKILL
                        txtQuestTask.text = txtQuestTask.text & TmpStr & "Gain level " & .amount & " " & GetSkillName(.MainData)
                End Select
            End With
        Next I
    End If
    
    For I = TID To Quest(Index).CLI(CID).Max_Actions
        With Quest(Index).CLI(CID).Action(I)
            If Not .ActionID >= 1 Or Not .ActionID <= 4 Then Exit For
            II = II + 1
            If II > 1 Then TmpStr = vbNewLine Else TmpStr = vbNullString
            Select Case .ActionID
                Case TASK_GATHER
                    txtQuestTask.text = txtQuestTask.text & TmpStr & "Gather " & .amount & " " & Trim$(Item(.MainData).Name)
                Case TASK_KILL
                    txtQuestTask.text = txtQuestTask.text & TmpStr & "Kill " & .amount & " " & Trim$(NPC(.MainData).Name)
                Case TASK_GETSKILL
                    txtQuestTask.text = txtQuestTask.text & TmpStr & "Gain level " & .amount & " " & GetSkillName(.MainData)
            End Select
        End With
    Next I
    
    If NPC(Quest(Index).CLI(CID).ItemIndex).ShowQuestCompleteIcon = 1 Then
        txtQuestTask.text = "Task(s) complete. Go back and speak with " & Trim$(NPC(Quest(Index).CLI(CID).ItemIndex).Name)
    End If
End Sub

Private Sub lstTitles_Click()
    Dim I As Byte
    
    ' Check if we're setting it to one we already have as our current title
    If lstTitles.ListIndex = Player(MyIndex).CurTitle Then Exit Sub
        
    If Not lstTitles.ListIndex = 0 Then
        For I = 1 To MAX_TITLES
            If Not Player(MyIndex).CurTitle = I Then
                If lstTitles.List(lstTitles.ListIndex) = Trim$(title(I).Name) Then
                    lblDesc.Caption = Trim$(title(I).Desc)
                    Call SendSetTitle(I)
                    Exit For
                End If
            End If
        Next
    Else
        lblDesc.Caption = "None."
        Call SendSetTitle(0)
    End If
End Sub

Private Sub lstTitles_GotFocus()
    SetGameFocus
End Sub

Private Sub lstFoes_GotFocus()
    SetGameFocus
End Sub

Private Sub lstFriends_GotFocus()
    SetGameFocus
End Sub

Private Sub lstGuild_GotFocus()
    SetGameFocus
End Sub

Private Sub mapPreviewSwitch_Click()
    If mapPreviewSwitch.Value Then
        mapPreviewSwitch.Picture = LoadResPicture("MAP_DOWN", vbResBitmap)
        frmMapPreview.Show
    Else
        mapPreviewSwitch.Picture = LoadResPicture("MAP_UP", vbResBitmap)
        Unload frmMapPreview
    End If
End Sub

Private Sub picChatbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearButtons
    ResetOptionButtons
End Sub

Private Sub picEquipFace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub picOptionBlood_Click()
    If Options.Blood = 0 Then
        Options.Blood = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Blood = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionBlood, OptionButtons.Opt_Blood, Options.Blood)
End Sub

Private Sub picOptionDebug_Click()
    If Options.Debug = 0 Then
        Options.Debug = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Debug = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionDebug, OptionButtons.Opt_Debug, Options.Debug)
End Sub

Private Sub picOptionSwearFilter_Click()
    If Options.SwearFilter = 0 Then
        Options.SwearFilter = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.SwearFilter = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionSwearFilter, OptionButtons.Opt_SwearFilter, Options.SwearFilter)
End Sub

Private Sub picOptionSound_Click()
    If Options.Sound = 0 Then
        Options.Sound = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Sound = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionSound, OptionButtons.Opt_Sound, Options.Sound)
End Sub

Private Sub picOptionMouse_Click()
    If Options.Mouse = 0 Then
        Options.Mouse = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Mouse = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    MouseX = -1
    MouseY = -1
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionMouse, OptionButtons.Opt_Mouse, Options.Mouse)
End Sub

Private Sub picOptionMusic_Click()
    If Options.Music = 0 Then
        Options.Music = 1
        Call Audio.PlaySound(ButtonClick)
        
        ' Start playing music
        PlayMapMusic
    Else
        Options.Music = 0
        Call Audio.PlaySound(ButtonBuzzer)
        
        ' Stop playing music
        Audio.StopMusic
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionMusic, OptionButtons.Opt_Music, Options.Music)
End Sub

Private Sub picOptionWeather_Click()
    If Options.Weather = 0 Then
        Options.Weather = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Weather = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionWeather, OptionButtons.Opt_Weather, Options.Weather)
End Sub

Private Sub picOptionAutoTile_Click()
    If Options.Autotile = 0 Then
        Options.Autotile = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Autotile = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionAutoTile, OptionButtons.Opt_AutoTile, Options.Autotile)
End Sub

Private Sub picOptionBattleMusic_Click()
    If Options.BattleMusic = 0 Then
        Options.BattleMusic = 1
        Call Audio.PlaySound(ButtonClick)
        
        ' Start playing music
        PlayMapMusic
    Else
        Options.BattleMusic = 0
        Call Audio.PlaySound(ButtonBuzzer)
        If Trim$(Map.Music) = vbNullString Then
            Call Audio.StopMusic
        Else
            Call Audio.PlayMusic(Trim$(Map.Music))
        End If
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionBattleMusic, OptionButtons.Opt_BattleMusic, Options.BattleMusic)
End Sub

Private Sub picOptionTitle_Click()
    If Options.Titles = 0 Then
        Options.Titles = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Titles = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionTitle, OptionButtons.Opt_Title, Options.Titles)
End Sub

Private Sub picOptionPlayerVitals_Click()
    If Options.PlayerVitals = 0 Then
        Options.PlayerVitals = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.PlayerVitals = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionPlayerVitals, OptionButtons.Opt_PlayerVitals, Options.PlayerVitals)
End Sub

Private Sub picOptionNPCVitals_Click()
    If Options.NPCVitals = 0 Then
        Options.NPCVitals = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.NPCVitals = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionNpcVitals, OptionButtons.Opt_NPCVitals, Options.NPCVitals)
End Sub

Private Sub picOptionLevel_Click()
    If Options.Levels = 0 Then
        Options.Levels = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Levels = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionLevel, OptionButtons.Opt_Level, Options.Levels)
End Sub

Private Sub picOptionGuild_Click()
    If Options.Guilds = 0 Then
        Options.Guilds = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Guilds = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionGuild, OptionButtons.Opt_Guilds, Options.Guilds)
End Sub

Private Sub picOptionWASD_Click()
    If Options.WASD = 0 Then
        Options.WASD = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.WASD = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionWASD, OptionButtons.Opt_WASD, Options.WASD)
End Sub

Private Sub picOptionBlood_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Blood)
    If OptionButton(OptionButtons.Opt_Blood).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionBlood, OptionButtons.Opt_Blood, 2 + Options.Blood)
End Sub

Private Sub picOptionDebug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Debug)
    If OptionButton(OptionButtons.Opt_Debug).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionDebug, OptionButtons.Opt_Debug, 2 + Options.Debug)
End Sub

Private Sub picOptionSwearFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_SwearFilter)
    If OptionButton(OptionButtons.Opt_SwearFilter).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionSwearFilter, OptionButtons.Opt_SwearFilter, 2 + Options.SwearFilter)
End Sub

Private Sub picOptionSound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Sound)
    If OptionButton(OptionButtons.Opt_Sound).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionSound, OptionButtons.Opt_Sound, 2 + Options.Sound)
End Sub

Private Sub picOptionMouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Mouse)
    If OptionButton(OptionButtons.Opt_Mouse).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionMouse, OptionButtons.Opt_Mouse, 2 + Options.Mouse)
End Sub

Private Sub picOptionMusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Music)
    If OptionButton(OptionButtons.Opt_Music).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionMusic, OptionButtons.Opt_Music, 2 + Options.Music)
End Sub

Private Sub picOptionWeather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Weather)
    If OptionButton(OptionButtons.Opt_Weather).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionWeather, OptionButtons.Opt_Weather, 2 + Options.Weather)
End Sub

Private Sub picOptionBattleMusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_BattleMusic)
    If OptionButton(OptionButtons.Opt_BattleMusic).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionBattleMusic, OptionButtons.Opt_BattleMusic, 2 + Options.BattleMusic)
End Sub

Private Sub picOptionTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Title)
    If OptionButton(OptionButtons.Opt_Title).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionTitle, OptionButtons.Opt_Title, 2 + Options.Titles)
End Sub

Private Sub picOptionPlayerVitals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_PlayerVitals)
    If OptionButton(OptionButtons.Opt_PlayerVitals).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionPlayerVitals, OptionButtons.Opt_PlayerVitals, 2 + Options.PlayerVitals)
End Sub

Private Sub picOptionNPCVitals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_NPCVitals)
    If OptionButton(OptionButtons.Opt_NPCVitals).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionNpcVitals, OptionButtons.Opt_NPCVitals, 2 + Options.NPCVitals)
End Sub

Private Sub picOptionLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Level)
    If OptionButton(OptionButtons.Opt_Level).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionLevel, OptionButtons.Opt_Level, 2 + Options.Levels)
End Sub

Private Sub picOptionGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Guilds)
    If OptionButton(OptionButtons.Opt_Guilds).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionGuild, OptionButtons.Opt_Guilds, 2 + Options.Guilds)
End Sub

Private Sub picOptionWASD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_WASD)
    If OptionButton(OptionButtons.Opt_WASD).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionWASD, OptionButtons.Opt_WASD, 2 + Options.WASD)
End Sub

Private Sub picOptionAutoTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_AutoTile)
    If OptionButton(OptionButtons.Opt_AutoTile).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionAutoTile, OptionButtons.Opt_AutoTile, 2 + Options.Autotile)
End Sub

Private Sub picEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picEventChat_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ToggleChatLock(Optional ByVal ForceLock As Boolean = False, Optional ByVal SoundEffect As Boolean = True)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ForceLock Then
        ChatLocked = True
    Else
        ChatLocked = Not ChatLocked
    End If
    
    If ChatLocked Then
        If SoundEffect Then Call Audio.PlaySound(ButtonBuzzer)
        frmMain.txtMyChat.text = vbNullString
        frmMain.txtMyChat.Enabled = False
        Exit Sub
    Else
        If SoundEffect Then Call Audio.PlaySound(ButtonClick)
        frmMain.txtMyChat.Enabled = True
         frmMain.txtMyChat.SetFocus
    End If
    
    Call SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ToggleChatLock", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

'--------------------------------------------------------------------------------
' Project    :       Client
' Procedure  :       picButton_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       SETH-PC
' Date-Time  :       5/28/2015-2:51:00 PM
'
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub picButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not CurButton_Main = Index Then
        Call Audio.PlaySound(ButtonClick)
        
        ' Don't set it if it's the trade/GUI adjusting button
        If Not Index = 5 And Not Index = 14 And Not Index = 15 Then
            CurButton_Main = Index
            picButton(Index).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(Index).FileName & "_click.jpg")
            Call ResetMainButtons
        End If
        
        Call TogglePanel(Index)
    Else ' Hide the panel, if it is the open one
        CurButton_Main = 0
        LastButton_Main = 0
        Call ResetMainButtons
        Call Audio.PlaySound(ButtonClick)
        Call TogglePanel(0)
    End If
    SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not LastButton_Main = Index And Not CurButton_Main = Index Then
        Call ResetMainButtons
        picButton(Index).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(Index).FileName & "_hover.jpg")
        Call Audio.PlaySound(ButtonHover)
        LastButton_Main = Index
    End If
    Call ClearChatButton(0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub TogglePanel(ByVal PanelNum As Long)
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Don't close panels if it's the trade button
    If Not PanelNum = 5 Then
        Call CloseAllPanels
    End If
    
    Select Case PanelNum
        Case 1
            picInventory.Visible = True
            picInventory.ZOrder (0)
        Case 2
            picSpells.Visible = True
            picSpells.ZOrder (0)
        Case 3
            picCharacter.Visible = True
            picCharacter.ZOrder (0)
        Case 4
            picOptions.Visible = True
            picOptions.ZOrder (0)
        Case 5
            If MyTargetType = TARGET_TYPE_PLAYER And Not MyTarget = MyIndex Then
                SendTradeRequest
            Else
                AddText "Invalid trade target.", BrightRed
            End If
        Case 6
            picParty.Visible = True
            picParty.ZOrder (0)
        Case 7
            picFriends.Visible = True
            picFriends.ZOrder (0)
        Case 8
            If GetPlayerGuild(MyIndex) = vbNullString Then
                picGuild_No.Visible = True
                picGuild_No.ZOrder (0)
            Else
                picGuild.Visible = True
                picGuild.ZOrder (0)
            End If
        Case 9
            For I = 1 To Skills.Skill_Count - 1
                lblSkill.Item(I - 1).Caption = GetSkillName(I)
                lblLevel.Item(I - 1).Caption = Player(MyIndex).Skills(I).Level
                lblSkillExp.Item(I - 1).Caption = Player(MyIndex).Skills(I).exp & "/" & GetPlayerNextSkillLevel(MyIndex, I)
            Next
            picSkills.Visible = True
            picSkills.ZOrder (0)
        Case 10
            picTitles.Visible = True
            picTitles.ZOrder (0)
        Case 11
            picQuest.Visible = True
            picQuest.ZOrder (0)
            Call LoadQuests
        Case 12
            picFoes.Visible = True
            picFoes.ZOrder (0)
        Case 14
            ButtonsVisible = Not ButtonsVisible
            If ButtonsVisible Then
                MainButton(14).FileName = "btn_hidepanels"
            Else
                MainButton(14).FileName = "btn_showpanels"
            End If
            Call ResetMainButtons
            Call ToggleButtons(ButtonsVisible)
        Case 15
            GUIVisible = Not GUIVisible
            If GUIVisible Then
                MainButton(15).FileName = "btn_hidegui"
            Else
                MainButton(15).FileName = "btn_showgui"
            End If
            Call ResetMainButtons
            Call ToggleGUI(GUIVisible)
        Case 16
            picEquipment.Visible = True
            picEquipment.ZOrder (0)
    End Select
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "TogglePanel", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResetMainButtons()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_MAINBUTTONS
        If Not CurButton_Main = I Then
            picButton(I).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(I).FileName & "_norm.jpg")
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ResetMainButtons", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picForm_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picForm_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picFriends_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picFriends_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picGuild_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long, rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Hotbar
    For I = 1 To MAX_HOTBAR
        With rec_pos
            .Top = picHotbar.Top - picHotbar.Top
            .Left = picHotbar.Left - picHotbar.Left + (HotbarOffsetX * (I - 1)) + (32 * (I - 1))
            .Right = .Left + 32
            .Bottom = picHotbar.Top - picHotbar.Top + 32
        End With
        
        If X >= rec_pos.Left And X <= rec_pos.Right Then
            If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                SendSwapHotbarSlots DragHotbarSlot, I
            End If
        End If
    Next
    
    DragHotbarSlot = 0
    picTempInv.Visible = False
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picHotbar_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long, I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum > 0 Then
        If Button = 1 Then
            If ShiftDown Then
                DragHotbarSlot = SlotNum
                
                For I = 1 To MAX_PLAYER_SPELLS
                    If Hotbar(DragHotbarSlot).Slot = PlayerSpells(I) Then
                        DragHotbarSpell = I
                    End If
                Next
            Else
                SendHotbarUse SlotNum
            End If
        ElseIf Button = 2 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If DragHotbarSlot > 0 Then
        If Hotbar(DragHotbarSlot).sType = 1 Then
            Call DrawDraggedItem(X + picHotbar.Left - 16, Y + picHotbar.Top - 16, True)
        Else
            Call DrawDraggedSpell(X + picHotbar.Left - 16, Y + picHotbar.Top - 16, True)
        End If
        picSpellDesc.Visible = False
        picItemDesc.Visible = False
        LastSpellDesc = 0 ' No spell was last loaded
        LastItemDesc = 0 ' No item was last loaded
        Exit Sub
    Else
        SlotNum = IsHotbarSlot(X, Y)
        
        If SlotNum <> 0 Then
              If Hotbar(SlotNum).sType = 1 Then ' Item
                X = X + picHotbar.Left - picItemDesc.Width - 1
                Y = Y + picHotbar.Top
                UpdateItemDescWindow Hotbar(SlotNum).Slot, X, Y
                LastItemDesc = Hotbar(SlotNum).Slot ' Set it so you don't re-set values
                Exit Sub
              ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
                X = X + picHotbar.Left - picSpellDesc.Width - 1
                Y = Y + picHotbar.Top
                UpdateSpellDescWindow Hotbar(SlotNum).Slot, X, Y
                LastSpellDesc = Hotbar(SlotNum).Slot

                For I = 1 To MAX_PLAYER_SPELLS
                    If Hotbar(SlotNum).Slot = PlayerSpells(I) Then
                        LastSpellSlotDesc = I
                    End If
                Next
                Exit Sub
              End If
          End If
    End If
    
    Call ClearChatButton(0)
    ClearButtons
    picSpellDesc.Visible = False
    picItemDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    picTempInv.Visible = False
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picOptions_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picParty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picParty_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picPet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picPet_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picScreen_DblClick()
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    MouseX = -1
    MouseY = -1
    
    ' Mouse
    If CurX = GetPlayerX(MyIndex) And CurY = GetPlayerY(MyIndex) Then
        Call CheckMapGetItem
    End If
    Exit Sub
   
' Error Handler
ErrorHandler:
    HandleError "Form_DblClick", "frmMain", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InMapEditor Then
        If chkEyeDropper.Value = 1 And displayTilesets = False Then
            Call MapEditorEyeDropper
        Else
            If displayTilesets And Not (X < 0 Or Y < 0 Or _
            X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width Or _
            Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height) And Button = 1 Then
                Call MapEditorChooseTile(Button, X, Y)
            ElseIf ControlDown And Button = 1 Then
                MapEditorFillSelection
                Exit Sub
            ElseIf ControlDown And Button = 2 Then
                MapEditorClearSelection
                Exit Sub
            ElseIf ShiftDown And Button = 1 Then
                MapEditorEyeDropper
                Exit Sub
            ElseIf Button = vbMiddleButton Then
                If frmMain.chkTilesets.Value Then
                    chkTilesets.Value = 0
                Else
                    chkTilesets.Value = 1
                End If
            ElseIf Button = vbRightButton Then
                If ShiftDown Then
                    ' Admin warp if we're pressing shift and right clicking
                    If GetPlayerAccess(MyIndex) >= STAFF_MAPPER Then
                        If CanMoveNow Then
                            AdminWarp CurX, CurY
                        End If
                    End If
                ElseIf InMapEditor And frmEditor_Map.OptEvents.Value Then
                    DeleteEvent CurX, CurY
                End If
            End If
            If Not displayTilesets Then
                Call MapEditorMouseDown(Button, X, Y, False)
                redrawMapCache = True
            End If
        End If
    Else
        ' Left click
        If Button = vbLeftButton Then
            If Options.Mouse = 1 Then
                ' Mouse
                If CurX = GetPlayerX(MyIndex) And CurY = GetPlayerY(MyIndex) Then
                    Call CheckMapGetItem
                Else
                    MouseX = CurX
                    MouseY = CurY
                End If
            End If
            
            ' Right click
        ElseIf Button = vbRightButton Then
        
            If ShiftDown Then
                ' Admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= STAFF_MAPPER Then
                    If CanMoveNow Then
                        AdminWarp CurX, CurY
                    End If
                End If
            Else
                Call PlayerSearch(CurX, CurY)
            End If
        End If
    End If

    Call SetGameFocus
    frmMain.picSpellDesc.Visible = False
    frmMain.picItemDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
    ' Error handler
ErrorHandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
    
    If InMapEditor Then
        If displayTilesets Then
            If frmEditor_Map.scrlAutotile.Value = 0 Then
                Call frmEditor_Map.MapEditorDrag(Button, X, Y)
            End If
        Else
            Call MapEditorMouseDown(Button, X, Y, False)
        End If
        
        If (LastX <> CurX Or LastY <> CurY) And frmEditor_Map.chkRandom.Value = 0 And Button >= 1 And Not displayTilesets Then
            redrawMapCache = True
        End If
    ElseIf Button = vbLeftButton And Options.Mouse = 1 Then
        ' Mouse
        If CurX = GetPlayerX(MyIndex) And CurY = GetPlayerY(MyIndex) Then
            Call CheckMapGetItem
        Else
            MouseX = CurX
            MouseY = CurY
        End If
    End If
    
    LastX = CurX
    LastY = CurY
    
    ' Set the description windows off
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    LastButton_Main = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_TRADES
        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            With TempRec
                .Top = ShopTop + ((ShopOffsetY + PIC_Y) * ((I - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + PIC_X) * (((I - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    
    ' Reset all buttons
    Call ResetMainButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picShop_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShopItem As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ShopItem = IsShopItem(X, Y)
    
    If ShopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(ShopItem)
                    If .CostItem > 0 And .CostItem2 = 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", BrightGreen
                    ElseIf .CostItem2 > 0 And .CostItem = 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", BrightGreen
                    ElseIf .CostItem > 0 And .CostItem2 > 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & " and " & .CostValue2 & " " & Trim$(Item(.CostItem2).Name) & ".", BrightGreen
                    Else
                        Exit Sub
                    End If
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem ShopItem
        End Select
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_dblClick()
    Dim ShopItem As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ShopItem = IsShopItem(ShopX, ShopY)
    
    If ShopItem > 0 Then
        BuyItem ShopItem
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picShopItems_dblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShopSlot As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ShopX = X
    ShopY = Y
    
    ShopSlot = IsShopItem(X, Y)

    If ShopSlot <> 0 Then
        X2 = X + picShop.Left + picShopItems.Left + 4
        Y2 = Y + picShop.Top + picShopItems.Top + 12
        UpdateItemDescWindow Shop(InShop).TradeItem(ShopSlot).Item, X2, Y2
        LastItemDesc = Shop(InShop).TradeItem(ShopSlot).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_DblClick()
    Dim SpellNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InTrade > 0 Or InBank Or InShop > 0 Or InChat Then Exit Sub

    SpellNum = IsPlayerSpell(SpellX, SpellY)

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        Call CastSpell(SpellNum)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SpellSlot As Byte
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellX = X
    SpellY = Y
    
    SpellSlot = IsPlayerSpell(X, Y)
    
    If DragSpellSlot > 0 Then
        Call DrawDraggedSpell(X + picSpells.Left - 16, Y + picSpells.Top - 16)
    Else
        If SpellSlot <> 0 Then
            X2 = picSpells.Left - picSpellDesc.Width - 4
            Y2 = picSpells.Top + 12
            UpdateSpellDescWindow PlayerSpells(SpellSlot), X2, Y2
            LastSpellDesc = PlayerSpells(SpellSlot)
            LastSpellSlotDesc = SpellSlot
            Exit Sub
        End If
    End If
    
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If DragSpellSlot > 0 Then
        ' Drag and Drop
        For I = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + PIC_X) * (((I - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If Not DragSpellSlot = I Then
                        If Not DialogueIndex = DIALOGUE_TYPE_FORGET Then
                            SendChangeSpellSlots DragSpellSlot, I
                        End If
                        Exit For
                    End If
                End If
            End If
        Next
        
        ' Hotbar
        For I = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picSpells.Top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (I - 1)) + (32 * (I - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picSpells.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpellSlot, I
                    DragSpellSlot = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpellSlot = 0
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SpellNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picSpellDesc.Visible = False
    LastSpellDesc = 0
    SpellNum = IsPlayerSpell(SpellX, SpellY)
    
    If Button = 1 Then ' left click
        If SpellNum <> 0 Then
            DragSpellSlot = SpellNum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' Right click
        If SpellNum > 0 And SpellNum <= MAX_PLAYER_SPELLS Then
            Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(SpellNum)).Name) & "?", DIALOGUE_TYPE_FORGET, True, SpellNum
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picToggleButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetMainButtons
End Sub

Private Sub picTitles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picTitles_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateItemDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num) ' Set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateItemDescWindow TradeTheirOffer(TradeNum).num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).num ' Set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GettingMap Then Exit Sub

    ' Set focus if making it visible
    If KeyAscii = vbKeyReturn Then
        If picEventChat.Visible Then
            If frmMain.lblEventChatContinue.Visible Then
                frmMain.lblEventChatContinue_Click
                KeyAscii = 0
                Exit Sub
            End If
        End If
        
        If picChatbox.Visible Then
            If txtMyChat.text = vbNullString Then
                If picCurrency.Visible = False Then
                    If picDialogue.Visible = False Then
                        Call ToggleChatLock
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    Call HandleKeyPresses(KeyAscii)

    ' Check if we need to call a label
    If frmMain.picCurrency.Visible Then
        If KeyAscii = vbKeyReturn Then Call lblCurrencyOk_Click
        If KeyAscii = vbKeyEscape Then Call lblCurrencyCancel_Click
    End If
    
    If frmMain.picDialogue.Visible Then
        If lblDialogue_Button(1).Visible Then
            If KeyAscii = vbKeyReturn Then Call lblDialogue_Button_Click(1)
        Else
            If KeyAscii = vbKeyReturn Then Call lblDialogue_Button_Click(2)
            If KeyAscii = vbKeyEscape Then Call lblDialogue_Button_Click(3)
        End If
    End If
    
    ' Prevents textbox on error ding soundnly be assigned to
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or ControlDown Then KeyAscii = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Make sure they can't press keys until they are in the game
    If InGame = False Then Exit Sub

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access >= STAFF_MODERATOR Then
                If FormVisible("frmAdmin") Then
                    If GetForegroundWindow = frmAdmin.hWnd Then
                        Unload frmAdmin
                    ElseIf GetForegroundWindow <> frmAdmin.hWnd Then
                        BringWindowToTop (frmAdmin.hWnd)
                    End If
                Else
                    InitAdminPanel
                End If
            End If
            
    End Select
    
    If ChatLocked Then
        If TempPlayer(MyIndex).Moving = NO Then
            If KeyCode = vbKeyHome Then
                If Player(MyIndex).Dir <> DIR_UP Then
                    Call SetPlayerDir(MyIndex, DIR_UP)
    
                    If Last_Dir <> GetPlayerDir(MyIndex) Then
                        Call SendPlayerDir
                        Last_Dir = GetPlayerDir(MyIndex)
                    End If
                End If
            End If
    
            If KeyCode = vbKeyEnd Then
                If Player(MyIndex).Dir <> DIR_DOWN Then
                    Call SetPlayerDir(MyIndex, DIR_DOWN)
    
                    If Last_Dir <> GetPlayerDir(MyIndex) Then
                        Call SendPlayerDir
                        Last_Dir = GetPlayerDir(MyIndex)
                    End If
                End If
            End If
    
            If KeyCode = vbKeyDelete And Not InMapEditor Then
                If Player(MyIndex).Dir <> DIR_LEFT Then
                    Call SetPlayerDir(MyIndex, DIR_LEFT)
    
                    If Last_Dir <> GetPlayerDir(MyIndex) Then
                        Call SendPlayerDir
                        Last_Dir = GetPlayerDir(MyIndex)
                    End If
                End If
            End If
    
            If KeyCode = vbKeyPageDown Then
                If Player(MyIndex).Dir <> DIR_RIGHT Then
                    Call SetPlayerDir(MyIndex, DIR_RIGHT)
    
                    If Last_Dir <> GetPlayerDir(MyIndex) Then
                        Call SendPlayerDir
                        Last_Dir = GetPlayerDir(MyIndex)
                    End If
                End If
            End If
        End If
    End If
    
    ' Handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' Handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If
    
    If ChatLocked Then
         ' Hotbar
        If frmMain.picCurrency.Visible Or (Not ChatLocked And Options.WASD = 1) Then
            ' Do nothing
        Else
            If Options.WASD = 1 Then
                For I = 1 To MAX_HOTBAR - 3 '
                    If KeyCode = 48 + I Or KeyCode = 96 + I Then
                        SendHotbarUse I
                    End If
                Next
                ' Hot bar button 0
                If KeyCode = 48 Or KeyCode = 96 Then SendHotbarUse 10
                
                ' Hot bar button -
                If KeyCode = 189 Or KeyCode = 109 Then SendHotbarUse 11
                
                ' Hot bar button +
                If KeyCode = 187 Or KeyCode = 107 Then SendHotbarUse 12
                Exit Sub
            Else
                For I = 1 To MAX_HOTBAR
                    If KeyCode = 111 + I Then
                        SendHotbarUse I
                    End If
                Next
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDialogue_Change()
    If DialogueIndex = DIALOGUE_TYPE_CHANGEGUILDACCESS Then
        If Not txtDialogue.text = vbNullString Then
            If Not IsNumeric(txtDialogue.text) Then txtDialogue.text = 1
            If txtDialogue.text < 1 Then txtDialogue.text = 1
            If txtDialogue.text > MAX_GUILDACCESS Then txtDialogue.text = MAX_GUILDACCESS
        End If
    End If
End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MyText = txtMyChat
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtChat_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub lblUseItem_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call UseItem
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblUseItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim Value As Long
    Dim Multiplier As Double
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
        ' Are we in a shop
        If InShop > 0 Then
            If Not TryingToFixItem Then
                SellItem InvNum
            Else
                FixItem InvNum
                TryingToFixItem = False
            End If
            Exit Sub
        End If
        
        ' In Bank
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
                If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                    CurrencyMenu = 2 ' Deposit
                    lblCurrency.Caption = "How many do you want to deposit?"
                    TmpCurrencyItem = InvNum
                    txtCurrency.text = vbNullString
                    picCurrency.Visible = True
                    picCurrency.ZOrder (0)
                    txtCurrency.SetFocus
                Else
                    Call DepositItem(InvNum, 1)
                End If
            Else
                Call DepositItem(InvNum, 0)
            End If
            Exit Sub
        End If
        
        ' In trade
        If InTrade > 0 Then
            ' Exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).num = InvNum Then
                    ' Is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).stackable = 1 Then
                        ' Only exit out if we're offering all of it
                        If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
                If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                    CurrencyMenu = 4 ' Offer in trade
                    lblCurrency.Caption = "How many do you want to trade?"
                    TmpCurrencyItem = InvNum
                    txtCurrency.text = vbNullString
                    picCurrency.Visible = True
                    picCurrency.ZOrder (0)
                    txtCurrency.SetFocus
                Else
                    Call TradeItem(InvNum, 1)
                End If
            Else
                Call TradeItem(InvNum, 0)
            End If
            Exit Sub
        End If
        
        ' Don't use an item if it is None or Auto Life
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_AUTOLIFE Then
            AddText "You can't use this type of item!", BrightRed
            Exit Sub
        End If
        
        ' Reset Stat Points
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_RESETSTATS Then
            Dialogue "Reset Stats", "Are you sure you wish to reset your stats?", DIALOGUE_TYPE_RESETSTATS, True, InvNum
            Exit Sub
        End If
        
        ' Use item if not doing anything else
        Call SendUseItem(InvNum)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(MyIndex, I) > 0 And GetPlayerEquipment(MyIndex, I) <= MAX_ITEMS Then
            With TempRec
                .Top = EquipSlotTop(I)
                .Bottom = .Top + PIC_Y
                .Left = EquipSlotLeft(I)
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsEqItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
            With TempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsInvItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(I) > 0 And PlayerSpells(I) <= MAX_PLAYER_SPELLS Then
            With TempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + PIC_X) * (((I - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsPlayerSpell = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim TempRec As RECT
    Dim I As Long
    Dim ItemNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_INV
        If Yours Then
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
        Else
            ItemNum = TradeTheirOffer(I).num
        End If

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            With TempRec
                .Top = InvTop - 12 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsTradeItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InTrade > 0 Then Exit Sub
    
    InvNum = IsInvItem(X, Y)
    
    If Button = 1 Then
        If InvNum > 0 And InvNum <= MAX_INV Then
            DragInvSlot = InvNum
            Exit Sub
        End If
    ElseIf Button = 2 Then
        If InvNum > 0 And InvNum <= MAX_INV Then
            Call DropItem(InvNum)
        End If
    End If

    SetGameFocus
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Byte
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    InvX = X
    InvY = Y

    If DragInvSlot > 0 Then
        If InTrade > 0 Then Exit Sub
        Call DrawDraggedItem(X + picInventory.Left - 16, Y + picInventory.Top - 16)
    Else
        InvNum = IsInvItem(X, Y)

        If Not InvNum = 0 Then
            ' Exit out if we're offering that item
            If InTrade > 0 Then
                For I = 1 To MAX_INV
                    If TradeYourOffer(I).num = InvNum Then
                        ' Is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).stackable = 1 Then
                            ' Only exit out if we're offering all of it
                            If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            
            X = picInventory.Left - picItemDesc.Width - 4
            Y = picInventory.Top + 12
            UpdateItemDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' Set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InTrade > 0 Then Exit Sub
    
    If DragInvSlot > 0 Then
        ' Drag and Drop
        For I = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If Not DragInvSlot = I Then
                        SendChangeInvSlots DragInvSlot, I
                        Exit For
                    End If
                End If
            End If
        Next
        
        ' Hotbar
        For I = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picInventory.Top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (I - 1)) + (32 * (I - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picInventory.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlot, I
                    DragInvSlot = 0
                    picTempInv.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlot = 0
    picTempInv.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' *****************
' ** Char Window **
' *****************
Private Sub picEquipment_Click()
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    EqNum = IsEqItem(EqX, EqY)

    If Not EqNum = 0 Then
        SendUnequip EqNum
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picEquipment_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picEquipment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If Not EqNum = 0 Then
        X = X + picEquipment.Left - picItemDesc.Width + 12
        Y = Y + picEquipment.Top - picItemDesc.Height + 16
        UpdateItemDescWindow GetPlayerEquipment(MyIndex, EqNum), X, Y
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' Set it so you don't re-set values
        Exit Sub
    End If
    
    Call ClearChatButton(0)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    ClearButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picEquipment_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Bank
Private Sub picBank_DblClick()
    Dim BankNum As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    DragBankSlot = 0

    BankNum = IsBankItem(BankX, BankY)
    
    If Not BankNum = 0 Then
        If Item(GetBankItemNum(BankNum)).stackable = 1 Then
            If GetBankItemValue(BankNum) > 1 Then
                CurrencyMenu = 3 ' Withdraw
                lblCurrency.Caption = "How many do you want to withdraw?"
                TmpCurrencyItem = BankNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder (0)
                txtCurrency.SetFocus
                Exit Sub
            Else
                WithdrawItem BankNum, 1
                Exit Sub
            End If
        Else
            WithdrawItem BankNum, 1
            Exit Sub
        End If
        WithdrawItem BankNum, 0
        Exit Sub
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BankNum As Long
                        
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    BankNum = IsBankItem(X, Y)
    
    If Not BankNum = 0 Then
        If Button = 1 Then
            DragBankSlot = BankNum
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If DragBankSlot > 0 Then
        For I = 1 To MAX_BANK
            With rec_pos
                .Top = BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + PIC_X) * (((I - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragBankSlot <> I Then
                        SwapBankSlots DragBankSlot, I
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlot = 0
    picTempBank.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BankNum As Long, ItemNum As Long, ItemType As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    BankX = X
    BankY = Y
    
    If DragBankSlot > 0 Then
        Call DrawBankItem(X + picBank.Left, Y + picBank.Top)
    Else
        BankNum = IsBankItem(X, Y)
        
        If BankNum <> 0 Then
            X2 = X + picBank.Left + 4
            Y2 = Y + picBank.Top + 4
            UpdateItemDescWindow Bank.Item(BankNum).num, X2, Y2
            LastItemDesc = Bank.Item(BankNum).num
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsBankItem = 0
    
    For I = 1 To MAX_BANK
        If GetBankItemNum(I) > 0 And GetBankItemNum(I) <= MAX_ITEMS Then
            With TempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + PIC_X) * (((I - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsBankItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub txtTransChat_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call SetGameFocus
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtTransChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CloseAllPanels()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    picInventory.Visible = False
    picSpells.Visible = False
    picCharacter.Visible = False
    picOptions.Visible = False
    picGuild.Visible = False
    picGuild_No.Visible = False
    picFriends.Visible = False
    picParty.Visible = False
    picEquipment.Visible = False
    picFoes.Visible = False
    picTitles.Visible = False
    picSkills.Visible = False
    picQuest.Visible = False
    picQuestDesc.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CloseAllPanels", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DropItem(ByVal InvNum As Byte)
    If InvNum > 0 And InvNum <= MAX_INV Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
            If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                CurrencyMenu = 1 ' drop
                lblCurrency.Caption = "How many do you want to drop?"
                TmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder (0)
                txtCurrency.SetFocus
                Exit Sub
            Else
                Call SendDropItem(InvNum, 1)
            End If
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
End Sub

Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
    Select Case uMsg
        Case WM_ACTIVATEAPP
            taskBarClick = True
        Case WM_MOVE
                If lng_hWnd = hwndLastActiveWnd Then
                    Dim rectt As modAdvMapEditor.RECTTT
                    GetWindowSize lng_hWnd, rectt
                    If FormVisible("frmAdmin") And adminMin Then
                        frmAdmin.centerMiniVert PixelsToTwips((rectt.Right - rectt.Left), 0), PixelsToTwips((rectt.Bottom - rectt.Top), 1), PixelsToTwips(rectt.Left, 0), PixelsToTwips(rectt.Top, 1)
                    End If
                    If FormVisible("frmMapPreview") And lng_hWnd = frmMain.hWnd Then
                        frmMapPreview.Move frmMain.Left - frmMapPreview.Width, frmMain.Top
                    End If
                    If FormVisible("frmMapPreview") Then
                        frmEditor_Map.Move frmMain.Left - frmEditor_Map.Width - 136, frmMain.Top + frmMapPreview.Height
                    Else
                        frmEditor_Map.Move frmMain.Left - frmEditor_Map.Width - 136, frmMain.Top
                    End If
                End If
        Case WM_NCACTIVATE
            If wParam Then
                hwndLastActiveWnd = lng_hWnd
                    Dim rectt2 As modAdvMapEditor.RECTTT
                    GetWindowSize lng_hWnd, rectt2
                    If FormVisible("frmAdmin") And adminMin Then
                        frmAdmin.centerMiniVert PixelsToTwips((rectt2.Right - rectt2.Left), 0), PixelsToTwips((rectt2.Bottom - rectt2.Top), 1), PixelsToTwips(rectt2.Left, 0), PixelsToTwips(rectt2.Top, 1)
                    End If
            End If
        Case WM_LBUTTONDOWN
            MainLButtonDown lng_hWnd
        Case WM_LBUTTONUP
            MainLButtonUp lng_hWnd
        Case WM_CAPTURECHANGED
            MainCaptureChanged lng_hWnd, lParam
        Case WM_MOUSEMOVE
            MainMouseMove lng_hWnd
            If InMapEditor Then
                If GetForegroundWindow = hWnd Or GetForegroundWindow = picScreen.hWnd Then
                    picScreen.SetFocus
                End If
            End If
        Case WM_GETMINMAXINFO 'Prevent Resizing, so we can keep nice frame when turning off CAPTION.
            If Not taskBarClick Then
                MainPreventResizing Me.hWnd, (Me.Width \ Screen.TwipsPerPixelX), (Me.Height \ Screen.TwipsPerPixelY), lParam
            Else
                taskBarClick = False
            End If
        Case WM_SETFOCUS
            If lng_hWnd = mapPreviewSwitch.hWnd Or lng_hWnd = chkEyeDropper.hWnd Or lng_hWnd = cmdSave.hWnd Or lng_hWnd = cmdRevert.hWnd Or lng_hWnd = cmdDelete.hWnd Or lng_hWnd = cmdProperties.hWnd Then
                bHandled = True
                lReturn = 1
            End If
        Case WM_MOUSEWHEEL
            If InMapEditor Then
                Dim Up As Boolean, curTil As Long
                Up = IIf(HiWord(wParam) = 120, False, True)
                If displayTilesets And chkLayers.Value = 0 And layersActive = False Then
                    curTil = frmEditor_Map.scrlTileSet.Value
                    frmEditor_Map.scrlTileSet.Value = (IIf((curTil = 1 And Not Up) Or (curTil = NumTileSets And Up), curTil, IIf(Up, 1, -1) + curTil))
                    'lblTitle = "UBER Map Editor - " & "Tileset: " & frmEditor_Map.scrlTileSet.Value
                Else
                    getCurrentMapLayerName
                    frmEditor_Map.optLayer(IIf((currentMapLayerNum = 1 And Not Up) Or (currentMapLayerNum = Layer_Count - 1 And Up), currentMapLayerNum, IIf(Up, 1, -1) + currentMapLayerNum)).Value = 1
                    getCurrentMapLayerName
                    displayTilesets = False
                    chkLayers.Value = 0
                End If

            End If
    End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
End Sub

Public Sub LoadQuests()
Dim I As Long
Dim Cnt As Long
Dim SEP_CHAR As String * 1
    frmMain.lstQuests.Clear
    Cnt = 1
    For I = 1 To MAX_QUESTS
        If Not InStr(Trim$(Quest(I).Name), SEP_CHAR) > 0 Then
            frmMain.lstQuests.AddItem Cnt & ": " & Trim$(Quest(I).Name)
            Cnt = Cnt + 1
        End If
    Next I
End Sub

Public Function FindQuest(ByVal QuestName As String) As Long
Dim I As Long
    QuestName = Mid$(QuestName, 4, Len(QuestName))

    For I = 1 To MAX_QUESTS
        If LCase$(Trim$(Quest(I).Name)) = LCase$(Trim$(QuestName)) Then
            FindQuest = I
            Exit Function
        End If
    Next I
End Function
