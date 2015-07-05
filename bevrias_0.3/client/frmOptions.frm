VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Options"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.Frame Frame4 
         Caption         =   "Startup Music"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   3135
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   980
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Ignore Titile Music"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Play Title Music"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "File name including .mid:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Item Name"
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3135
         Begin VB.OptionButton Option6 
            Caption         =   "Show Item Name When END is pressed down"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2895
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Don't Show Item Name On Map"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   2895
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Showing Item Name On Map"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   4320
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   3105
         TabIndex        =   6
         Top             =   240
         Width           =   3135
         Visible         =   0   'False
         Begin VB.CommandButton Command2 
            Caption         =   "Close"
            Height          =   255
            Left            =   2280
            TabIndex        =   8
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   $"frmOptions.frx":0000
            Height          =   1095
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "FPS Speed"
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "Help"
            Height          =   255
            Left            =   2400
            TabIndex        =   5
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Fast FPS (For Great Computers)"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   2895
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Normal FPS (Default)"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   2895
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Bad FPS (For Slow Computers)"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2895
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
End Sub

Private Sub Command3_Click()
frmOptions.Visible = False
End Sub

Private Sub Command4_Click()
WriteINI "CONFIG", "MusicName", Text1.Text, App.Path & "\config.ini"
End Sub

Private Sub Form_Load()
Text1.Text = ReadINI("CONFIG", "MusicTitle", App.Path & "\config.ini")
If ReadINI("CONFIG", "TitleMusic", App.Path & "\config.ini") = 0 Then
Option8.Value = True
End If
If ReadINI("CONFIG", "TitleMusic", App.Path & "\config.ini") = 1 Then
Option7.Value = True
End If
If ReadINI("ITEMS", "ItemName", App.Path & "\config.ini") = 1 Then
Option4.Value = True
End If
If ReadINI("ITEMS", "ItemName", App.Path & "\config.ini") = 0 Then
Option5.Value = True
End If
If ReadINI("ITEMS", "ItemName", App.Path & "\config.ini") = 2 Then
Option6.Value = True
End If
If ReadINI("GAMESPEED", "FPS", App.Path & "\config.ini") = 0 Then
Option1.Value = True
End If
If ReadINI("GAMESPEED", "FPS", App.Path & "\config.ini") = 1 Then
Option2.Value = True
End If
If ReadINI("GAMESPEED", "FPS", App.Path & "\config.ini") = 2 Then
Option3.Value = True
End If
End Sub

Private Sub Option1_Click()
WriteINI "GAMESPEED", "FPS", 0, App.Path & "\config.ini"
End Sub

Private Sub Option2_Click()
WriteINI "GAMESPEED", "FPS", 1, App.Path & "\config.ini"
End Sub

Private Sub Option3_Click()
WriteINI "GAMESPEED", "FPS", 2, App.Path & "\config.ini"
End Sub

Private Sub Option4_Click()
WriteINI "ITEMS", "ItemName", 1, App.Path & "\config.ini"
End Sub

Private Sub Option5_Click()
WriteINI "ITEMS", "ItemName", 0, App.Path & "\config.ini"
End Sub

Private Sub Option6_Click()
WriteINI "ITEMS", "ItemName", 2, App.Path & "\config.ini"
End Sub

Private Sub Option7_Click()
WriteINI "CONFIG", "TitleMusic", 1, App.Path & "\config.ini"
End Sub

Private Sub Option8_Click()
WriteINI "CONFIG", "TitleMusic", 0, App.Path & "\config.ini"
End Sub
