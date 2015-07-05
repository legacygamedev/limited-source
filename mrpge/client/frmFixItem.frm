VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Fix Item)"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFixItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmFixItem.frx":0442
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmFixItem.frx":25B5
      Top             =   120
      Width           =   195
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmFixItem.frx":475B
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image chkFix 
      Height          =   480
      Left            =   480
      Picture         =   "frmFixItem.frx":4C3B
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmFixItem.frx":509F
      Top             =   0
      Width           =   7680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   6195
      Left            =   0
      Picture         =   "frmFixItem.frx":7C6E
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFix_Click()
    Call SendData("fixitem" & SEP_CHAR & cmbItem.ListIndex + 1 & SEP_CHAR & END_CHAR)
End Sub




Private Sub Image3_Click()
Call GameDestroy
End Sub

Private Sub Image4_Click()
frmFixItem.WindowState = vbMinimized
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub
