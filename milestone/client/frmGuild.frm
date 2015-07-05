VERSION 5.00
Begin VB.Form frmGuild 
   BackColor       =   &H00789298&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Guild"
   ClientHeight    =   6000
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuild.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   330
      TabIndex        =   1
      Top             =   1620
      Width           =   2895
   End
   Begin VB.TextBox txtGuild 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   0
      Top             =   2220
      Width           =   2895
   End
   Begin VB.Label Command1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
      Height          =   210
      Left            =   2745
      TabIndex        =   4
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   3
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guild Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   330
      TabIndex        =   2
      Top             =   2025
      Width           =   825
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Packet As String
Packet = "MAKEGUILD" & SEP_CHAR & txtName.Text & SEP_CHAR & txtGuild.Text & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub
