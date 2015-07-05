VERSION 5.00
Begin VB.Form frmSign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign Attribute"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Set Sign Text"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
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
         TabIndex        =   5
         Top             =   2160
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
         Left            =   2520
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Line1 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox Line2 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Line3 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line2:"
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
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line3:"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    SignLine1 = Line1.Text
    SignLine2 = Line2.Text
    SignLine3 = Line3.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Line1.Text = SignLine1
    Line2.Text = SignLine2
    Line3.Text = SignLine3
End Sub
