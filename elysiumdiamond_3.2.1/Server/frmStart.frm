VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Starting Elysium Diamond 3 Account Editor"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7635
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Choose the account"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtChoose 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblChoose 
      Caption         =   "Type the account to edit"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Starting up..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
Dim I As Long

    'txtChoose
    
    If AccountExist(txtChoose.Text) = True Then
            
        Call LoadPlayer(1, txtChoose.Text)
        
        InitStart
        
        Me.Hide
        frmRun.Show
    Else
        MsgBox "Account doesn't exist, please type a correct name.", vbOKOnly, "Error"
        txtChoose.Text = vbNullString
    End If
    DoEvents

End Sub

Private Sub Form_Load()
    
    lblStatus.Caption = "Setting up account editor..."
    DoEvents
    
    Call InitEditor
    
    'Load frmRun
    'frmRun.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub
