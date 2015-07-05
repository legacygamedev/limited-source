VERSION 5.00
Begin VB.Form frmDataEditor 
   Caption         =   "rootSource Data Editor"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2760
      TabIndex        =   14
      Top             =   5520
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame framGame 
      Caption         =   "Game Settings"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   3735
      Begin VB.TextBox txtSite 
         Height          =   285
         Left            =   240
         MaxLength       =   500
         TabIndex        =   12
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtGameName 
         Height          =   285
         Left            =   240
         MaxLength       =   255
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Website"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Name"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraMusic 
      Caption         =   "Music Settings"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3735
      Begin VB.ComboBox cmbExt 
         Height          =   315
         ItemData        =   "frmDataEditor.frx":0000
         Left            =   240
         List            =   "frmDataEditor.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music Extension"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame frmConnections 
      Caption         =   "Connection Settings"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmDataEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim FileName As String
    Dim F As Long
    
    FileName = App.Path & "\data\config.dat"

    With GameData
    
        .IP = Trim$(txtIP.Text)
        .Port = Val(txtPort.Text)
        .MusicExt = Trim$(cmbExt.Text)
        .GameName = Trim$(txtGameName.Text)
        .WebAddress = Trim$(txtSite.Text)
    
    End With
    
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , GameData
    Close #F
    
    MsgBox "Data Saved!", vbOKOnly, "Information"
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim FileName As String
    Dim F As Long
    
    txtIP.MaxLength = NAME_LENGTH
    txtPort.MaxLength = NAME_LENGTH
    
    
    FileName = App.Path & "\data\config.dat"
    
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , GameData
    Close #F
    
    With GameData
    
        txtIP.Text = Trim$(.IP)
        txtPort.Text = Trim$(.Port)
        cmbExt.Text = Trim$(.MusicExt)
        txtGameName = Trim$(.GameName)
        txtSite.Text = Trim$(.WebAddress)
    
    End With
    
    
End Sub
