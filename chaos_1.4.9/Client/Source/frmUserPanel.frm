VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUserPanel 
   Caption         =   "User Panel"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   353
      TabMaxWidth     =   1587
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "frmUserPanel.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnclose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Commands"
      TabPicture(1)   =   "frmUserPanel.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Game Play"
      TabPicture(2)   =   "frmUserPanel.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame9"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Credits"
      TabPicture(3)   =   "frmUserPanel.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(1)=   "Label2"
      Tab(3).Control(2)=   "lblUserName"
      Tab(3).Control(3)=   "lblLevel"
      Tab(3).Control(4)=   "lblRegister"
      Tab(3).Control(5)=   "Frame13"
      Tab(3).Control(6)=   "Frame12"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Bugs"
      TabPicture(4)   =   "frmUserPanel.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "btnSend"
      Tab(4).Control(1)=   "btnCancel"
      Tab(4).Control(2)=   "txtBugReport"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton btnSend 
         Caption         =   "Send"
         Height          =   495
         Left            =   -74760
         TabIndex        =   41
         Top             =   1920
         Width           =   3855
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   -74760
         TabIndex        =   40
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   2175
         Begin VB.CommandButton cmdAlign 
            Caption         =   "View Alignment"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdwhosonlinelist 
            Caption         =   "Whos Online"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdFriend 
            Caption         =   "Friends List"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Suggestions"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton btnReport 
            Caption         =   "Report Bugs"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdWhosOnline 
            Caption         =   "Whos Online"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Check FPS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   21
         Top             =   435
         Width           =   2055
         Begin VB.CommandButton cmdTrade 
            Caption         =   "Start Trade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdAccpTrade 
            Caption         =   "Accept Trade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdDeclnTrade 
            Caption         =   "Decline Trade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Chat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2400
         TabIndex        =   16
         Top             =   435
         Width           =   2055
         Begin VB.TextBox txtChat1 
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdChat 
            Caption         =   "Start Chat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdDeclnChat 
            Caption         =   "Decline Chat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Accept Chat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnclose 
         Caption         =   "Close User Panel"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Caption         =   "Game Commands"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -74760
         TabIndex        =   13
         Top             =   555
         Width           =   4215
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Text            =   "frmUserPanel.frx":008C
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74760
         TabIndex        =   11
         Top             =   555
         Width           =   2055
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Text            =   "frmUserPanel.frx":03D1
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "More Buttons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   9
         Top             =   2235
         Width           =   4215
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "frmUserPanel.frx":042D
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Ask For Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -72600
         TabIndex        =   7
         Top             =   555
         Width           =   2055
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "frmUserPanel.frx":0578
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Chaos Engine Creator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   5
         Top             =   555
         Width           =   2295
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "frmUserPanel.frx":0636
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Chaos Engine History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -72480
         TabIndex        =   3
         Top             =   555
         Width           =   1935
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "frmUserPanel.frx":06A8
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Exit Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         TabIndex        =   1
         Top             =   1995
         Width           =   2055
         Begin VB.CommandButton Command13 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
      End
      Begin RichTextLib.RichTextBox txtBugReport 
         Height          =   1545
         Left            =   -74760
         TabIndex        =   42
         Top             =   360
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   2725
         _Version        =   393217
         BackColor       =   16777215
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmUserPanel.frx":06F5
      End
      Begin VB.Label lblRegister 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74880
         TabIndex        =   30
         Top             =   5475
         Width           =   1860
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Player Level"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   3675
         Width           =   2835
      End
      Begin VB.Label lblUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Player Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   3435
         Width           =   2835
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chaos Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   -73800
         TabIndex        =   27
         Top             =   3075
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74760
         TabIndex        =   26
         Top             =   3075
         Width           =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   4440
         Y1              =   5715
         Y2              =   5715
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   555
         Y2              =   3795
      End
   End
End
Attribute VB_Name = "frmUserPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
txtBugReport.Text = ""
End Sub

Private Sub btnSend_Click()
If Not txtBugReport.Text = "" Then
        Call SendBugReport(Trim(txtBugReport.Text))
        txtBugReport.Text = ""
    End If
End Sub

Private Sub cmdAlign_Click()
frmStats.Show
Unload Me
End Sub

Private Sub cmdDeclnChat_Click()
Call SendData("dchat" & SEP_CHAR & END_CHAR)
End Sub

Private Sub cmdFriend_Click()
If frmMirage.picFriend.Visible = False Then
frmMirage.picFriend.Visible = True
Else
frmMirage.picFriend.Visible = False
End If
End Sub

Private Sub cmdGuild_Click()
End Sub

Private Sub cmdRefresh_Click()
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
End Sub

Private Sub cmdTrade_Click()
Dim TradeName
TradeName = Text1.Text
Call SendTradeRequest(TradeName)
End Sub

Private Sub cmdAccpTrade_Click()
Call SendAcceptTrade
MyText = ""
End Sub

Private Sub cmdChat_Click()
Call SendData("playerchat" & SEP_CHAR & Trim(txtChat1.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub cmdDeclnTrade_Click()
Call SendDeclineTrade
MyText = ""
End Sub

Private Sub btnclose_Click()
frmUserPanel.Visible = False
End Sub

Private Sub cmdwhosonlinelist_Click()
If frmMirage.picWhosOnline.Visible = False Then
Call SendOnlineList
    frmMirage.picWhosOnline.Visible = True
    Else
    frmMirage.picWhosOnline.Visible = False
    End If
End Sub

Private Sub Command1_Click()
Call SendData("achat" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command13_Click()
Call GameDestroy
End Sub

Private Sub cmdWhosOnline_Click()
Call SendWhosOnline
MyText = ""
End Sub

Private Sub Command2_Click()
frmSuggestions.Show vbModal
Unload Me
End Sub

Private Sub Command3_Click()
frmStats.Show
End Sub

Private Sub Command4_Click()
Call AddText("FPS: " & GameFPS, Yellow)
End Sub

Private Sub Form_Load()
'MsgBox " DO NOT SPAM The Bug Report Section OR YOU WILL BE BANNED!"
End Sub
