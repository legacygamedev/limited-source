VERSION 5.00
Begin VB.Form frmTraining 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3900
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
   Icon            =   "frmTraining.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTraining.frx":08CA
   ScaleHeight     =   277.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStat 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "frmTraining.frx":47066
      Left            =   480
      List            =   "frmTraining.frx":47076
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblTrain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      UseMnemonic     =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   5040
      Width           =   1155
   End
End
Attribute VB_Name = "frmTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = GAME_NAME
    'Me.Picture = LoadPicture(App.Path & "\gfx\interface\Menu.bmp")
    cmbStat.ListIndex = 0
End Sub

Private Sub lblTrain_Click()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 6
    Buffer.WriteInteger CUseStatPoint
    Buffer.WriteLong cmbStat.ListIndex
    Call SendData(Buffer.ToArray())
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

