VERSION 5.00
Begin VB.Form frmSpriteChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Change Attribute"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmSpriteChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Set Sprite"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1200
         Width           =   480
         Begin VB.PictureBox picSprites 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   5
         Top             =   600
         Width           =   3855
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
         TabIndex        =   4
         Top             =   2760
         Width           =   1935
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
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   30000
         TabIndex        =   2
         Top             =   2400
         Width           =   4335
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Note: In order for the sprite change to display on the screen, the map name must end with *"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   4080
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   11
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   465
         TabIndex        =   10
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   9
         Top             =   2160
         Width           =   405
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   525
         TabIndex        =   7
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   6
         Top             =   1560
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmSpriteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

    If SpritePic < scrlSprite.min Then
        SpritePic = scrlSprite.min
    End If
    scrlSprite.Value = SpritePic
    If SpriteItem < scrlItem.min Then
        SpriteItem = scrlItem.min
    End If
    scrlItem.Value = SpriteItem
    If SpritePrice < scrlCost.min Then
        SpritePrice = scrlCost.min
    End If
    scrlCost.Value = SpritePrice

    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * 64) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * PIC_Y) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * PIC_Y, SRCCOPY)
    End If

End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = scrlCost.Value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.Value = 0 Then
        lblItem.Caption = "No Cost"
        Exit Sub
    Else
        lblItem.Caption = scrlItem.Value & " - " & Trim$(Item(scrlItem.Value).name)
    End If

    If Item(scrlItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlCost.Enabled = True
    Else
        scrlCost.Enabled = False
    End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = scrlSprite.Value
    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * 64) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * PIC_Y) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
End Sub

Private Sub Timer1_Timer()
    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * 64) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.Top = (scrlSprite.Value * PIC_Y) * -1
    ' Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
End Sub
