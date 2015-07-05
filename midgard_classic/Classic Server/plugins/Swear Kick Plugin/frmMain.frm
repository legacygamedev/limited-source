VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swear Kick Plugin By Pc"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraKicked 
      Caption         =   "Kicked"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3495
      Begin VB.ListBox lstKicked 
         Height          =   1425
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame fraIndex 
      Caption         =   "Bot Index"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Label lblIndex 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
