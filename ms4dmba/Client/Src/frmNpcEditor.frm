VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDrop 
      Caption         =   "Drop Item Data"
      Height          =   2655
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Width           =   4935
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   3000
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   32
         Top             =   1680
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   31
         Top             =   2160
         Value           =   1
         Width           =   3255
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   3000
         TabIndex        =   30
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Drop Item Chance 1 out of"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4320
         TabIndex        =   40
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Num"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblItem 
         Caption         =   "Item"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblItemName 
         Height          =   375
         Left            =   840
         TabIndex        =   37
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Value"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4320
         TabIndex        =   35
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Index"
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtAttackSay 
      Height          =   345
      Left            =   960
      TabIndex        =   27
      Top             =   600
      Width           =   3975
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
      TabIndex        =   25
      Top             =   9360
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   6360
      Width           =   1455
   End
   Begin VB.HScrollBar scrlMagic 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   20
      Top             =   2400
      Width           =   2895
   End
   Begin VB.HScrollBar scrlSpeed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   17
      Top             =   2160
      Width           =   2895
   End
   Begin VB.HScrollBar scrlDefense 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   14
      Top             =   1920
      Width           =   2895
   End
   Begin VB.HScrollBar scrlStrength 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   11
      Top             =   1680
      Width           =   2895
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   8
      Top             =   1440
      Value           =   1
      Width           =   2895
   End
   Begin VB.ComboBox cmbBehavior 
      Height          =   360
      ItemData        =   "frmNpcEditor.frx":0000
      Left            =   1440
      List            =   "frmNpcEditor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.HScrollBar scrlSprite 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
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
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label13 
      Caption         =   "Start HP"
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblStartHP 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   1320
      TabIndex        =   44
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Exp Given"
      Height          =   375
      Left            =   2640
      TabIndex        =   43
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblExpGiven 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3840
      TabIndex        =   42
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblSay 
      Caption         =   "Say"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Magic"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblMagic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Speed"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Defense"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDefense 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Strength"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblStrength 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Range"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Behavior"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Sprite"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblSprite 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
    scrlSprite.Max = NumSprites
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call NpcEditorOk
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    frmIndex.Show
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = CStr(scrlSprite.Value)
    
    Call NpcEditorBltSprite
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.Value)
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = CStr(scrlStrength.Value)
    lblStartHP.Caption = CStr(scrlStrength.Value * scrlDefense.Value)
    lblExpGiven.Caption = CStr(scrlStrength.Value * scrlDefense.Value * 2)
End Sub

Private Sub scrlDefense_Change()
    lblDefense.Caption = CStr(scrlDefense.Value)
    lblStartHP.Caption = CStr(scrlStrength.Value * scrlDefense.Value)
    lblExpGiven.Caption = CStr(scrlStrength.Value * scrlDefense.Value * 2)
End Sub

Private Sub scrlSpeed_Change()
    lblSpeed.Caption = CStr(scrlSpeed.Value)
    lblExpGiven.Caption = CStr(scrlStrength.Value * scrlDefense.Value * 2)
End Sub

Private Sub scrlMagic_Change()
    lblMagic.Caption = CStr(scrlMagic.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = CStr(scrlNum.Value)
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = CStr(scrlValue.Value)
End Sub

