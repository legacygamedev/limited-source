VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
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
   Icon            =   "frmNpcEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   43
      Top             =   6720
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info."
      TabPicture(0)   =   "frmNpcEditor.frx":2372
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStartHP"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpGiven"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDEF"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblSTR"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlDEF"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlSTR"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   4935
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   36
            Top             =   600
            Value           =   1
            Width           =   3375
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   35
            Top             =   1080
            Value           =   1
            Width           =   3375
         End
         Begin VB.TextBox txtChance 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   30
            Text            =   "0"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtSpawnSecs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   29
            Text            =   "0"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblNum 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   42
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Num:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Item:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblItemName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   39
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Value:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Drop Item Chance 1 out of:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Spawn Rate (in seconds):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   2040
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4935
         Begin VB.ComboBox cmbBehavior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmNpcEditor.frx":238E
            Left            =   960
            List            =   "frmNpcEditor.frx":23A1
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1800
            Width           =   3855
         End
         Begin VB.PictureBox picSprite 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4320
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   27
            Top             =   1200
            Width           =   480
         End
         Begin VB.HScrollBar scrlSprite 
            Height          =   255
            Left            =   960
            Max             =   255
            TabIndex        =   22
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   21
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtAttackSay 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   20
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Behaviour:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblSprite 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   26
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sprite:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Spoken:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   12
         Top             =   2640
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   11
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label lblSTR 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Strength:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblDEF 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Defence:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblExpGiven 
         Alignment       =   1  'Right Justify
         Caption         =   "EXP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label lblStartHP 
         Alignment       =   1  'Right Justify
         Caption         =   "HP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   6120
         Width           =   855
      End
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   9480
      Top             =   120
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   9360
      Width           =   480
   End
   Begin VB.HScrollBar scrlMAGI 
      Height          =   375
      Left            =   7905
      Max             =   255
      TabIndex        =   6
      Top             =   6195
      Width           =   2895
   End
   Begin VB.HScrollBar scrlSPEED 
      Height          =   375
      Left            =   7905
      Max             =   255
      TabIndex        =   3
      Top             =   5715
      Width           =   2895
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   375
      Left            =   7905
      Max             =   255
      TabIndex        =   0
      Top             =   4275
      Value           =   1
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "MAGI"
      Height          =   375
      Left            =   7065
      TabIndex        =   8
      Top             =   6195
      Width           =   855
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   10890
      TabIndex        =   7
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "SPD"
      Height          =   375
      Left            =   7065
      TabIndex        =   5
      Top             =   5715
      Width           =   855
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   10890
      TabIndex        =   4
      Top             =   5700
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Range"
      Height          =   375
      Left            =   7065
      TabIndex        =   2
      Top             =   4275
      Width           =   855
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   10890
      TabIndex        =   1
      Top             =   4260
      Width           =   495
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub scrlSprite_Change()
    lblSprite.Caption = STR(scrlSprite.Value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = STR(scrlSTR.Value)
    lblStartHP.Caption = STR(scrlSTR.Value * scrlDEF.Value)
    lblExpGiven.Caption = STR(scrlSTR.Value * scrlDEF.Value * 2)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = STR(scrlDEF.Value)
    lblStartHP.Caption = STR(scrlSTR.Value * scrlDEF.Value)
    lblExpGiven.Caption = STR(scrlSTR.Value * scrlDEF.Value * 2)
End Sub



Private Sub scrlMAGI_Change()
    lblMAGI.Caption = STR(scrlMAGI.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = STR(scrlNum.Value)
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim(Item(scrlNum.Value).Name)
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = STR(scrlValue.Value)
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

