VERSION 5.00
Begin VB.Form frmEditScript 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Scripts"
   ClientHeight    =   5775
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtMain 
      Height          =   5055
      Left            =   120
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmEditScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
