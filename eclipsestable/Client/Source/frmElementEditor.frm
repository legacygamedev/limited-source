VERSION 5.00
Begin VB.Form frmElementEditor 
   Caption         =   "Element Editor"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3375
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Edit Element"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
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
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrong 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   2
         Top             =   1200
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlWeak 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   1
         Top             =   1920
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         TabIndex        =   10
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strong:"
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
         TabIndex        =   9
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblStrong 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weak:"
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
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label lblWeak 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1680
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmElementEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ElementEditorOk
End Sub
Private Sub Command1_Click()
    Call ElementEditorCancel
End Sub

Private Sub Form_Load()
    scrlStrong.max = MAX_ELEMENTS
    scrlWeak.max = MAX_ELEMENTS
End Sub

Private Sub scrlStrong_Change()
    lblStrong.Caption = Element(scrlStrong.Value).name
End Sub

Private Sub scrlWeak_Change()
    lblWeak.Caption = Element(scrlWeak.Value).name
End Sub
