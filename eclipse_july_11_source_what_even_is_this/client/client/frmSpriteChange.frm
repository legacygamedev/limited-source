VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmSpriteChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Change Attribute"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmSpriteChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   -240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   -240
      Visible         =   0   'False
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
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
      TabCaption(0)   =   "Set Sprite"
      TabPicture(0)   =   "frmSpriteChange.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSprite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCost"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblItem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlSprite"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOk"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlCost"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlItem"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "picSprite"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   600
         Width           =   480
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   9
         Top             =   1200
         Width           =   4335
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   30000
         TabIndex        =   4
         Top             =   1800
         Width           =   4335
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
         Left            =   360
         TabIndex        =   3
         Top             =   2160
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
         Left            =   2520
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "No Cost"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   520
         TabIndex        =   10
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   465
         TabIndex        =   7
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sprite:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   470
         TabIndex        =   6
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmSpriteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmSpriteChange.Visible = False
End Sub

Private Sub cmdOk_Click()
    SpritePic = scrlSprite.Value
    SpriteItem = scrlItem.Value
    SpritePrice = scrlCost.Value
    scrlCost.Value = 0
    scrlSprite.Value = 0
    scrlItem.Value = 0
    frmSpriteChange.Visible = False
End Sub

Private Sub Form_Load()
    If SpritePic < scrlSprite.Min Then SpritePic = scrlSprite.Min
    scrlSprite.Value = SpritePic
    If SpriteItem < scrlItem.Min Then SpriteItem = scrlItem.Min
    scrlItem.Value = SpriteItem
    If SpritePrice < scrlCost.Min Then SpritePrice = scrlCost.Min
    scrlCost.Value = SpritePrice
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = scrlCost.Value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.Value = 0 Then
        lblItem.Caption = "No Cost"
        Exit Sub
    Else
        lblItem.Caption = scrlItem.Value & " - " & Trim(Item(scrlItem.Value).Name)
    End If

    If Item(scrlItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlCost.Enabled = True
    Else
        scrlCost.Enabled = False
    End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = scrlSprite.Value
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * PIC_Y, SRCCOPY)
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * PIC_Y, SRCCOPY)
End Sub
