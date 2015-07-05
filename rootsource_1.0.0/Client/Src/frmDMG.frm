VERSION 5.00
Begin VB.Form frmDMG 
   Caption         =   "Damage to Deal"
   ClientHeight    =   495
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnConfirm 
      Caption         =   "Confirm"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   10000
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmDMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConfirm_Click()
    EditorData1 = HScroll1.Value
    Unload Me
End Sub

Private Sub HScroll1_Change()
    Me.Caption = "Damage to Deal: " & HScroll1.Value
End Sub
