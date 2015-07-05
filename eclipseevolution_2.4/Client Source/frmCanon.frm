VERSION 5.00
Begin VB.Form frmCanon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canon Placement"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Canon"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   5055
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
         Left            =   840
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
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
         Left            =   3000
         TabIndex        =   8
         Top             =   2280
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   5
         Top             =   1200
         Width           =   5055
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5400
         TabIndex        =   11
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Animation: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5280
         TabIndex        =   6
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Damage: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5280
         TabIndex        =   3
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item used to fire:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmCanon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Me.Visible = False

End Sub

Private Sub cmdOk_Click()

    CanonItem = val(HScroll1.Value)
    CanonDamage = val(HScroll2.Value)
    '//!! Unknown controls
    'If optNorth.Value = True Then CanonDirection = 0
    'If optEast.Value = True Then CanonDirection = 1
    'If optSouth.Value = True Then CanonDirection = 2
    'If optWest.Value = True Then CanonDirection = 3
    Me.Visible = False

End Sub

Private Sub Form_Load()

    HScroll1.Max = MAX_ITEMS

End Sub

Private Sub HScroll1_Change()

    Label1.Caption = "Item used to fire: " & item(HScroll1.Value).Name
    Label2.Caption = val(HScroll1.Value)

End Sub

Private Sub HScroll2_Change()

    Label4.Caption = HScroll2.Value

End Sub

Private Sub HScroll3_Change()

    Label5.Caption = "Animation: " & Spell(HScroll3.Value).Name
    Label6.Caption = val(HScroll3.Value)

End Sub

