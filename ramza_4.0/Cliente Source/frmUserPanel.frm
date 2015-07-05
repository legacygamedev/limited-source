VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUserPanel 
   Caption         =   "Panel de Usuario"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
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
      Tab(0).Control(3)=   "btnclose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Comandos"
      TabPicture(1)   =   "frmUserPanel.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Jugabilidad"
      TabPicture(2)   =   "frmUserPanel.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame9"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Creditos"
      TabPicture(3)   =   "frmUserPanel.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label4"
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(3)=   "Label2"
      Tab(3).Control(4)=   "Label1"
      Tab(3).Control(5)=   "Frame13"
      Tab(3).Control(6)=   "Frame12"
      Tab(3).ControlCount=   7
      Begin VB.Frame Frame5 
         Caption         =   "Opciones"
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
         Left            =   2400
         TabIndex        =   29
         Top             =   2280
         Width           =   2055
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Actualizar"
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
            Left            =   360
            TabIndex        =   32
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Revisar FPS"
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
            Left            =   360
            TabIndex        =   31
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdWhosOnline 
            Caption         =   "Jugadores Online"
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
            Left            =   360
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comercio"
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
         TabIndex        =   24
         Top             =   720
         Width           =   2055
         Begin VB.CommandButton cmdTrade 
            Caption         =   "Empezar Comercio"
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
            TabIndex        =   28
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdAccpTrade 
            Caption         =   "Aceptar Comercio"
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
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdDeclnTrade 
            Caption         =   "Declinar Comercio"
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
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdParty 
            Caption         =   "Crear una Party"
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
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmdJoin 
            Caption         =   "Unirse a una Party"
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
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdLeave 
            Caption         =   "Dejar la Party"
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
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Charla"
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
         TabIndex        =   14
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtChat1 
            Height          =   285
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdChat 
            Caption         =   "Empezar Charla"
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
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdDeclnChat 
            Caption         =   "Declinar Charla*"
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
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Aceptar Charla*"
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
            TabIndex        =   15
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnclose 
         Caption         =   "Cerrar el Panel"
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
         TabIndex        =   13
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Caption         =   "Comandos del Juego"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -74760
         TabIndex        =   11
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
            Height          =   3495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "frmUserPanel.frx":0070
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Controles"
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
         TabIndex        =   9
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
            TabIndex        =   10
            Text            =   "frmUserPanel.frx":039B
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Mas botones"
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
         TabIndex        =   7
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
            TabIndex        =   8
            Text            =   "frmUserPanel.frx":03FD
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Ayuda"
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
         TabIndex        =   5
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
            TabIndex        =   6
            Text            =   "frmUserPanel.frx":053A
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Desarrolladores"
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
         TabIndex        =   3
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
            TabIndex        =   4
            Text            =   "frmUserPanel.frx":059D
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Mas Informacion"
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
         TabIndex        =   1
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
            TabIndex        =   2
            Text            =   "frmUserPanel.frx":0626
            Top             =   240
            Width           =   1695
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
         Caption         =   "Panel de Usuarios V 3.0"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "El sitio oficial de Ramza Engine:"
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
         TabIndex        =   37
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "http://www.ramzaengine.com.ar"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "El sitio oficial de la Comunidad Inovapc:"
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
         TabIndex        =   34
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "http://www.inovapc.net"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   3600
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

Private Sub cmdWhosOnline_Click()
Call SendWhosOnline
MyText = ""
End Sub

Private Sub Command4_Click()
Call AddText("FPS: " & GameFPS, Yellow)
End Sub

Private Sub Label2_Click()
Shell ("explorer http://www.ramzaengine.com.ar"), vbNormalNoFocus
End Sub

Private Sub Label3_Click()
Shell ("explorer http://www.google.com"), vbNormalNoFocus
End Sub

Private Sub Label6_Click()
Shell ("explorer http://www.inovapc.net"), vbNormalNoFocus
End Sub

