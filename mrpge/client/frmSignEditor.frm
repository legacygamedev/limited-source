VERSION 5.00
Begin VB.Form frmSignEditor 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtmsg 
      Height          =   2535
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmSignEditor.frx":0000
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtheader 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmSignEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
    Call SignEditorOk
End Sub

