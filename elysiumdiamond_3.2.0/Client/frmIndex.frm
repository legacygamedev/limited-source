VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Which Preference To Edit"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
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
      ItemData        =   "frmIndex.frx":0000
      Left            =   240
      List            =   "frmIndex.frx":0002
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
      TabPicture(0)   =   "frmIndex.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex + 1
    
    If InItemsEditor = True Then
        Call SendData(EDITITEM_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InNpcEditor = True Then
        Call SendData(EDITNPC_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InShopEditor = True Then
        Call SendData(EDITSHOP_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InSpellEditor = True Then
        Call SendData(EDITSPELL_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InEmoticonEditor = True Then
        Call SendData(EDITEMOTICON_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InArrowEditor = True Then
        Call SendData(EDITARROW_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    If InSpeechEditor = True Then
        Call SendData(EDITSPEECH_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    Unload frmIndex
End Sub

Private Sub cmdCancel_Click()
    InSpeechEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    Unload frmIndex
End Sub
