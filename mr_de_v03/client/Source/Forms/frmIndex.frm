VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Select"
      TabPicture(0)   =   "frmIndex.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstIndex"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         ItemData        =   "frmIndex.frx":0E5E
         Left            =   120
         List            =   "frmIndex.frx":0E60
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         TabIndex        =   1
         Top             =   2280
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex + 1
    
    If InItemsEditor Then SendEditItem
    If InNpcEditor Then SendEditNpc
    If InShopEditor Then SendEditShop
    If InSpellEditor Then SendEditSpell
    If InEmoticonEditor Then SendEditEmoticon
    If InAnimationEditor Then SendEditAnimation
    If InQuestEditor Then SendEditQuest
    Unload frmIndex
End Sub

Private Sub cmdCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InAnimationEditor = False
    InQuestEditor = False
    Unload frmIndex
End Sub

