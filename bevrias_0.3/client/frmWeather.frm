VERSION 5.00
Begin VB.Form frmWeather 
   Caption         =   "Edit Weather"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "General Properties"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "None"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rain"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Snow"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Thunder"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtint 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Set*"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Default*"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear*"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Day or Night"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblWeather 
         Caption         =   "Current Weather:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Weather Intensity:"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
lblWeather.Caption = "Current Weather: " & "None"
Call SendData("weather" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command2_Click()
lblWeather.Caption = "Current Weather: " & "Rain"
Call SendData("weather" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command3_Click()
lblWeather.Caption = "Current Weather: " & "Snow"
Call SendData("weather" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command4_Click()
lblWeather.Caption = "Current Weather: " & "Thunder"
Call SendData("weather" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command5_Click()
frmWeather.Visible = False
frmadmin.Visible = True
End Sub

Private Sub Command9_Click()
If GameTime = TIME_DAY Then
GameTime = TIME_NIGHT
Else
GameTime = TIME_DAY
End If
Call SendGameTime
MyText = ""
End Sub
