VERSION 5.00
Begin VB.Form frmRoof 
   BorderStyle     =   0  'None
   Caption         =   "Floor"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   1350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Housebox 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton optHouse1 
         Caption         =   "House 1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optHouse2 
         Caption         =   "House 2"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optHouse3 
         Caption         =   "House 3"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optHouse4 
         Caption         =   "House 4"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optHouse5 
         Caption         =   "House 5"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optHouse6 
         Caption         =   "House 6"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRoof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
