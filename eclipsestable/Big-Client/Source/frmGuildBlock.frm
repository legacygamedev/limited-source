VERSION 5.00
Begin VB.Form frmGuildBlock 
   Caption         =   "Guild Block"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "GuildBlock"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtGuild 
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
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Name to Pass Block:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmGuildBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    GuildBlock = txtGuild.Text
    Unload Me
End Sub
