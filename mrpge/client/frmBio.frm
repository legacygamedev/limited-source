VERSION 5.00
Begin VB.Form frmBio 
   Caption         =   "BIO"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtBio 
      Height          =   3045
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmBio.frx":0000
      Top             =   1440
      Width           =   4620
   End
   Begin VB.TextBox txtemail 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1980
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label Label3 
      Caption         =   "Bio"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "e-mail address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
    SendData ("SAVEBIO" & SEP_CHAR & Trim(txtBio) & SEP_CHAR & Trim(txtName) & SEP_CHAR & Trim(txtemail) & SEP_CHAR & END_CHAR)
    Unload Me
End Sub
