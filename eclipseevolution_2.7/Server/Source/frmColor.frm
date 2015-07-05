VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r As Integer
Dim g As Integer
Dim b As Integer

Private Sub Command1_Click()
    frmNews.Red = Slider1.value
    frmNews.Green = Slider2.value
    frmNews.Blue = Slider3.value
    frmNews.Text1.ForeColor = RGB(Slider1.value, Slider2.value, Slider3.value)
    frmNews.Text2.ForeColor = RGB(Slider1.value, Slider2.value, Slider3.value)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Slider1.value = frmNews.Red
    Slider2.value = frmNews.Green
    Slider3.value = frmNews.Blue
End Sub

Private Sub Slider1_Change()
    r = Slider1.value
    Shape1.BackColor = RGB(r, g, b)
End Sub
Private Sub Slider2_Change()
    g = Slider2.value
    Shape1.BackColor = RGB(r, g, b)
End Sub
Private Sub Slider3_Change()
    b = Slider3.value
    Shape1.BackColor = RGB(r, g, b)
End Sub

