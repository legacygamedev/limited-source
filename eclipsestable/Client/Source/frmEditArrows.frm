VERSION 5.00
Begin VB.Form frmEditArrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrow Editor"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "frmEditArrows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Edit Arrow"
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2520
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picArrows 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   0
               Width           =   480
            End
         End
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
         Left            =   1680
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
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
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   1215
      End
      Begin VB.HScrollBar scrlArrow 
         Height          =   255
         Left            =   240
         Max             =   500
         Min             =   1
         TabIndex        =   5
         Top             =   1200
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   4
         Top             =   2160
         Value           =   1
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
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
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   240
         Max             =   500
         Min             =   1
         TabIndex        =   2
         Top             =   2760
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range:"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label lblArrow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow:"
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
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
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
         TabIndex        =   8
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   45
   End
End
Attribute VB_Name = "frmEditArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ArrowEditorOk
End Sub

Private Sub Command1_Click()
    Call ArrowEditorCancel
End Sub

Private Sub scrlArrow_Change()
    lblArrow.Caption = "Arrow: " & scrlArrow.Value
    frmEditArrows.picArrows.Top = (scrlArrow.Value * 32) * -1
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.Value
End Sub
Private Sub scrlAmount_Change()
    lblAmount.Caption = "Amount: " & scrlAmount
End Sub

Private Sub SSTab1_DblClick()

End Sub
