VERSION 5.00
Begin VB.Form frmItemDesc 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   5
      Top             =   90
      Width           =   1530
   End
   Begin VB.Label lbl2Hand 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Top             =   1920
      Width           =   2235
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   1635
      Width           =   2235
   End
   Begin VB.Label lblStack 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   240
      Left            =   30
      TabIndex        =   2
      Top             =   1335
      Width           =   2235
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   240
      Left            =   30
      TabIndex        =   1
      Top             =   1005
      Width           =   2235
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ItemName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   2235
   End
End
Attribute VB_Name = "frmItemDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Move frmItemSpawner.Left - Width, frmItemSpawner.Top + frmItemSpawner.Height - Height
End Sub
