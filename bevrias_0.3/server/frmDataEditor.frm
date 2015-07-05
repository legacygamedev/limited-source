VERSION 5.00
Begin VB.Form frmDataEditor 
   Caption         =   "Data Editor"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1680
      TabIndex        =   35
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   1680
      TabIndex        =   33
      Top             =   3810
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save All"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   3518
      Width           =   3735
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label17 
      Caption         =   "Scrolling"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3820
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "Website"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3520
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Max Party Members"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Scripting"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Max Level"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Max Emoticons"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Max Guild Members"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Max Guilds"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Max Map Items"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Max Maps"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Max Spells"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Max Shops"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Max NPC's"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Max Items"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Max Players"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmDataEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call PutVar(App.Path & "\Data.ini", "CONFIG", "GameName", Text1.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "Max_Players", Text2.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "Max_Items", Text3.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "Max_NPCS", Text4.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS", Text5.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS", Text6.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS", Text7.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", Text8.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS", Text9.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", Text10.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", Text11.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL", Text12.text)
Call PutVar(App.Path & "\Data.ini", "CONFIG", "scripting", Text13.text)
Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS", Text14.text)
Call PutVar(App.Path & "\Data.ini", "CONFIG", "website", Text15.text)
Call PutVar(App.Path & "\Data.ini", "CONFIG", "port", Text16.text)
Call PutVar(App.Path & "\Data.ini", "CONFIG", "scrolling", Text17.text)
End Sub

Private Sub Command2_Click()
frmDataEditor.Visible = False
End Sub

Private Sub Form_Load()
    GAME_NAME = Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS")
    MAX_MAPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS")
    MAX_EMOTICONS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS")
    MAX_LEVEL = GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL")
    Scripting = GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting")
    MAX_PARTY_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS")
    Text1.text = GAME_NAME
    Text2.text = MAX_PLAYERS
    Text3.text = MAX_ITEMS
    Text4.text = MAX_NPCS
    Text5.text = MAX_SHOPS
    Text6.text = MAX_SPELLS
    Text7.text = MAX_MAPS
    Text8.text = MAX_MAP_ITEMS
    Text9.text = MAX_GUILDS
    Text10.text = MAX_GUILD_MEMBERS
    Text11.text = MAX_EMOTICONS
    Text12.text = MAX_LEVEL
    Text13.text = Scripting
    Text15.text = GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")
    Text16.text = GetVar(App.Path & "\Data.ini", "CONFIG", "Port")
    Text17.text = GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling")
    Text14.text = MAX_PARTY_MEMBERS
End Sub
