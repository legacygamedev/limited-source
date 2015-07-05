VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Custom Form"
   ClientHeight    =   480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   58
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustom 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCustom 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmNum As Long

Private Sub cmdCustom_Click(Index As Integer)
    Call SendData("CUSTOMFORMBUTTONCLICK" & SEP_CHAR & frmNum & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
End Sub

Private Sub Form_Load()
Dim I As Long

    For I = 1 To MAX_OBJECTS
        Load lblCustom(I)
        Load cmdCustom(I)
    Next I
End Sub

Private Sub lblCustom_Click(Index As Integer)
    Call SendData("CUSTOMFORMTEXTCLICK" & SEP_CHAR & frmNum & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
End Sub
