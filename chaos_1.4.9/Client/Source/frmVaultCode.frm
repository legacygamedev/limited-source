VERSION 5.00
Begin VB.Form frmVaultCode 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vault Password"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAuthenticate 
      Caption         =   "Authenticate"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtVaultCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   " ~Please Enter Your Vault Password~"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "*Vault Security*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Realms of The Wicked"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmVaultCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAuthenticate_Click()
Call SendVaultCode(txtVaultCode.Text)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

