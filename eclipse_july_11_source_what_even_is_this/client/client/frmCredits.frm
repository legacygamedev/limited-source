VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Engine Credits"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   3600
      X2              =   3600
      Y1              =   1440
      Y2              =   2640
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Eclipse Team Member's"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Malikona"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "AlanSpike"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Coke"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GSD"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Shannara"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Marsh"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pingu"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rafiki"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Globe"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "The Yellow Mole"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   2430
      TabIndex        =   0
      Top             =   5400
      Width           =   1320
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("GUI\Credits" & Ending) Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End Sub

