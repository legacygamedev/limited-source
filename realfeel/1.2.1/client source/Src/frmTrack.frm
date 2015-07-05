VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrack 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin Chat Device"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbChat 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssChat 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Full Chat"
      TabPicture(0)   =   "frmTrack.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtFullChat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Map Chat"
      TabPicture(1)   =   "frmTrack.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMapChat"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Broad Chat"
      TabPicture(2)   =   "frmTrack.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtBroadcastChat"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Global Chat"
      TabPicture(3)   =   "frmTrack.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtGlobalChat"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Private Chat"
      TabPicture(4)   =   "frmTrack.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtPrivateChat"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Admin Chat"
      TabPicture(5)   =   "frmTrack.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtAdminChat"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Tracker Chat"
      TabPicture(6)   =   "frmTrack.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtTrackerChat"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.TextBox txtPrivateChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtBroadcastChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtGlobalChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtAdminChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtTrackerChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtMapChat 
         Height          =   2415
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtFullChat 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
sbChat.SimpleText = "Admin Chat Device Loaded!"
End Sub

Private Sub Form_Terminate()
Call SendRemoveTracker(TrackName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SendRemoveTracker(TrackName)
End Sub
