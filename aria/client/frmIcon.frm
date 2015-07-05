VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniMap Icon Tile"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Set Icon"
      TabPicture(0)   =   "frmIcon.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "scrIcon"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "picIcon"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   840
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   5
         Top             =   390
         Width           =   120
      End
      Begin VB.HScrollBar scrIcon 
         Height          =   255
         Left            =   360
         Max             =   100
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Icon:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    IconNum = scrIcon.Value
    Unload Me
End Sub

Private Sub Form_Load()
    If IconNum < scrIcon.Min Then IconNum = scrIcon.Min
    scrIcon.Value = IconNum
    
    rec.Top = scrIcon.Value * 6
    rec.Bottom = (scrIcon.Value + 1) * 6
    rec.Left = 0
    rec.Right = 6
    
    rec_pos.Top = 0
    rec_pos.Left = 0
    rec_pos.Bottom = 6
    rec_pos.Right = 6
    
    Call DD_MiniMap.BltToDC(picIcon.hDC, rec, rec_pos)
End Sub

Private Sub scrIcon_Change()

    rec.Top = scrIcon.Value * 6
    rec.Bottom = (scrIcon.Value + 1) * 6
    rec.Left = 0
    rec.Right = 6
    
    rec_pos.Top = 0
    rec_pos.Left = 0
    rec_pos.Bottom = 6
    rec_pos.Right = 6
    
    Call DD_MiniMap.BltToDC(picIcon.hDC, rec, rec_pos)
End Sub

