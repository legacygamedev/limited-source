VERSION 5.00
Begin VB.Form frmMapSign 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtSign 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtHeader 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.HScrollBar scrlSign 
      Height          =   255
      Left            =   1440
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Sign"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMapSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
EditorSignNumber = scrlSign.Value
Unload Me
End Sub

Private Sub Form_Load()
scrlSign.Value = 1
End Sub

Private Sub scrlSign_Change()
If scrlSign.Value = 0 Then scrlSign.Value = 1
txtHeader.Text = Signs(scrlSign.Value).header
txtSign.Text = Signs(scrlSign.Value).msg
End Sub
