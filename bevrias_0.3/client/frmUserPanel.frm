VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUserPanel 
   Caption         =   "User Panel"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnclose"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Commands"
      TabPicture(1)   =   "frmUserPanel.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Game Play"
      TabPicture(2)   =   "frmUserPanel.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Credits"
      TabPicture(3)   =   "frmUserPanel.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label2"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame13"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame12"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
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
         Height          =   855
         Left            =   120
         TabIndex        =   34
         Top             =   4005
         Width           =   2055
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   240
            Width           =   855
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
            Left            =   120
            TabIndex        =   35
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
         TabIndex        =   29
         Top             =   720
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
            TabIndex        =   33
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   2055
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdParty 
            Caption         =   "Create Party"
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
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmdJoin 
            Caption         =   "Join Party"
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
            TabIndex        =   26
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdLeave 
            Caption         =   "Leave Party"
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
            Top             =   600
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
         TabIndex        =   19
         Top             =   405
         Width           =   2055
         Begin VB.TextBox txtChat1 
            Height          =   285
            Left            =   240
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdDeclnChat 
            Caption         =   "Decline Chat*"
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
            TabIndex        =   21
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Accept Chat*"
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
            TabIndex        =   20
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
         Left            =   2760
         TabIndex        =   18
         Top             =   4605
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
         TabIndex        =   16
         Top             =   360
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
            TabIndex        =   17
            Text            =   "frmUserPanel.frx":0070
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
         TabIndex        =   14
         Top             =   360
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
            TabIndex        =   15
            Text            =   "frmUserPanel.frx":06DC
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
         TabIndex        =   12
         Top             =   2040
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
            TabIndex        =   13
            Text            =   "frmUserPanel.frx":0738
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
         TabIndex        =   10
         Top             =   360
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
            TabIndex        =   11
            Text            =   "frmUserPanel.frx":08C7
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Engine Creator"
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
         TabIndex        =   8
         Top             =   360
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
            TabIndex        =   9
            Text            =   "frmUserPanel.frx":0946
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Engine History"
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
         TabIndex        =   6
         Top             =   360
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
            TabIndex        =   7
            Text            =   "frmUserPanel.frx":09D5
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Exit Game"
         Height          =   855
         Left            =   2400
         TabIndex        =   3
         Top             =   1965
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
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "*Save Game*"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Save Game"
         Height          =   975
         Left            =   2400
         TabIndex        =   1
         Top             =   2805
         Width           =   2055
         Begin VB.Label Label7 
            Caption         =   "You don't need to press Save. The Game will be saved if you press Exit."
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   4560
         Y1              =   3885
         Y2              =   3885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   525
         Y2              =   3765
      End
      Begin VB.Label Label5 
         Caption         =   "User Panel Version 0.3"
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
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Open the offical site for Bevrias Engine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   42
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "http://www.Bevrias.com"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "http://www.Bevrias-Engine.tk"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Open the offical forum for Bevrias Engine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "http://www.Ngage-Online.de/forums/_hosted_/..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   3480
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmUserPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdRefresh_Click()
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
End Sub

Private Sub cmdTrade_Click()
Call SendData("trade" & Text1.Text)
End Sub
Private Sub cmdJoin_Click()
Call SendJoinParty
End Sub

Private Sub cmdLeave_Click()
Call SendLeaveParty
End Sub

Private Sub cmdParty_Click()
Call SendPartyRequest("party" & Text2.Text)
MyText = ""
End Sub
Private Sub cmdAccpTrade_Click()
Call SendAcceptTrade
MyText = ""
End Sub

Private Sub cmdChat_Click()
Call SendData("playerchat" & txtChat1.Text)
MyText = ""
End Sub

Private Sub cmdDeclnTrade_Click()
Call SendDeclineTrade
MyText = ""
End Sub

Private Sub btnclose_Click()
frmUserPanel.Visible = False
End Sub

Private Sub Command13_Click()
Call GameDestroy
End Sub

Private Sub cmdWhosOnline_Click()
Call SendWhosOnline
MyText = ""
End Sub

Private Sub Command4_Click()
Call AddText("FPS: " & GameFPS, Yellow)
End Sub

Private Sub Label2_Click()
Shell ("explorer http://www.bevrias.com"), vbNormalNoFocus
End Sub

Private Sub Label3_Click()
Shell ("explorer http://www.Bevrias-Engine.tk"), vbNormalNoFocus
End Sub

Private Sub Label6_Click()
Shell ("explorer http://www.ngage-online.de/forums/_hosted_/bevrias-engine/index.php"), vbNormalNoFocus
End Sub

