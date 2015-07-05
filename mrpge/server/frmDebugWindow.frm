VERSION 5.00
Begin VB.Form frmDebugWindow 
   Caption         =   "editnpc"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      Height          =   6255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmDebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
'ShowDebug = False
End Sub

Private Sub txtDebug_Change()
On Error Resume Next
    txtDebug.SelStart = Len(txtDebug.Text) - 1
End Sub
