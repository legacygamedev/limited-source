VERSION 5.00
Begin VB.Form frmSPR 
   Caption         =   "Sprite to Change to:"
   ClientHeight    =   525
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmSPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    EditorData1 = HScroll1.Value
    Unload Me
End Sub

Private Sub HScroll1_Change()
    Me.Caption = "Sprite to Change to: " & HScroll1.Value
End Sub
