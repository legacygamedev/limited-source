VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Which Preference To Edit"
   ClientHeight    =   3495
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   5550
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
   ScaleHeight     =   3495
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
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
      TabIndex        =   1
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ListBox lstIndex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      ItemData        =   "frmIndex.frx":0CCA
      Left            =   240
      List            =   "frmIndex.frx":0CCC
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Choice"
      TabPicture(0)   =   "frmIndex.frx":0CCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
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

    If InItemsEditor Then
        Call SendData("edititem" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InNpcEditor Then
        Call SendData("editnpc" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InShopEditor Then
        Call SendData("editshop" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InElementEditor Then
        Call SendData("editelement" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InSpellEditor Then
        Call SendData("editspell" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InEmoticonEditor Then
        Call SendData("editemoticon" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InArrowEditor Then
        Call SendData("editarrow" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    Unload frmIndex
End Sub

Private Sub cmdCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InElementEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InArrowEditor = False

    Unload frmIndex
End Sub
