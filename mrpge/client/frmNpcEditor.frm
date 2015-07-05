VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmdType 
      Height          =   390
      ItemData        =   "frmNpcEditor.frx":0000
      Left            =   6585
      List            =   "frmNpcEditor.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   4455
      Width           =   3615
   End
   Begin VB.CheckBox chkBank 
      Caption         =   "Opens Bank"
      Height          =   300
      Left            =   8340
      TabIndex        =   57
      Top             =   3930
      Width           =   1770
   End
   Begin VB.CheckBox chkShop 
      Caption         =   "Opens Shop"
      Height          =   300
      Left            =   5505
      TabIndex        =   56
      Top             =   3960
      Width           =   1770
   End
   Begin VB.TextBox txtQuestNo 
      Height          =   390
      Left            =   7260
      TabIndex        =   55
      Text            =   "1"
      ToolTipText     =   "QuestID"
      Top             =   2175
      Width           =   1650
   End
   Begin VB.CheckBox chkQuest 
      Caption         =   "Starts a Quest:"
      Height          =   330
      Left            =   5370
      TabIndex        =   54
      Top             =   2220
      Width           =   1980
   End
   Begin VB.TextBox txtPoisonVital 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9495
      TabIndex        =   53
      Text            =   "20"
      ToolTipText     =   "Enter the hp the attack will remove"
      Top             =   1410
      Width           =   720
   End
   Begin VB.TextBox txtPoisonLength 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9495
      TabIndex        =   51
      Text            =   "20"
      ToolTipText     =   "Enter the number of time the poison will be effective for"
      Top             =   1740
      Width           =   720
   End
   Begin VB.CheckBox chkPoisonAttack 
      Caption         =   "Attacks with poison"
      Height          =   525
      Left            =   5355
      TabIndex        =   49
      Top             =   1470
      Width           =   1605
   End
   Begin VB.TextBox txtExpGiven 
      Height          =   390
      Left            =   2670
      TabIndex        =   47
      Text            =   "0"
      Top             =   4395
      Width           =   2355
   End
   Begin VB.TextBox txtStartHP 
      Height          =   390
      Left            =   225
      TabIndex        =   46
      Text            =   "0"
      Top             =   4365
      Width           =   2340
   End
   Begin VB.OptionButton optNoRespawn 
      Caption         =   "no"
      Height          =   375
      Left            =   8880
      TabIndex        =   45
      Top             =   1020
      Width           =   1215
   End
   Begin VB.OptionButton optYesRespawn 
      Caption         =   "Yes"
      Height          =   375
      Left            =   7560
      TabIndex        =   44
      Top             =   1035
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtSpawnSecs 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   8400
      TabIndex        =   42
      Text            =   "0"
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox txtAttackSay 
      Height          =   390
      Left            =   960
      TabIndex        =   40
      Top             =   510
      Width           =   3975
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   0
      Top             =   -45
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
      TabIndex        =   35
      Top             =   9360
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2670
      TabIndex        =   34
      Top             =   4740
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   150
      TabIndex        =   33
      Top             =   4740
      Width           =   2415
   End
   Begin VB.HScrollBar scrlValue 
      Height          =   375
      Left            =   6210
      Max             =   255
      TabIndex        =   31
      Top             =   3435
      Value           =   1
      Width           =   3375
   End
   Begin VB.HScrollBar scrlNum 
      Height          =   375
      Left            =   6210
      Max             =   1020
      TabIndex        =   25
      Top             =   2955
      Value           =   1
      Width           =   3375
   End
   Begin VB.TextBox txtChance 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   8400
      TabIndex        =   24
      Text            =   "0"
      Top             =   105
      Width           =   1815
   End
   Begin VB.HScrollBar scrlMAGI 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   20
      Top             =   2895
      Width           =   2895
   End
   Begin VB.HScrollBar scrlSPEED 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   17
      Top             =   2505
      Width           =   2895
   End
   Begin VB.HScrollBar scrlDEF 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   14
      Top             =   2115
      Width           =   2895
   End
   Begin VB.HScrollBar scrlSTR 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   11
      Top             =   1725
      Width           =   2895
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   8
      Top             =   1335
      Value           =   1
      Width           =   2895
   End
   Begin VB.ComboBox cmbBehavior 
      Height          =   390
      ItemData        =   "frmNpcEditor.frx":001E
      Left            =   1440
      List            =   "frmNpcEditor.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3315
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   75
      Width           =   3975
   End
   Begin VB.HScrollBar scrlSprite 
      Height          =   375
      Left            =   1080
      Max             =   255
      TabIndex        =   1
      Top             =   945
      Width           =   2895
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
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   945
      Width           =   480
   End
   Begin VB.Label Label20 
      Caption         =   "Behavior"
      Height          =   375
      Left            =   5415
      TabIndex        =   59
      Top             =   4470
      Width           =   1095
   End
   Begin VB.Line Line7 
      X1              =   5355
      X2              =   10155
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line6 
      X1              =   5385
      X2              =   10185
      Y1              =   2595
      Y2              =   2595
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Vital of poison"
      Height          =   300
      Left            =   7980
      TabIndex        =   52
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Length of poison"
      Height          =   300
      Left            =   7725
      TabIndex        =   50
      Top             =   1770
      Width           =   1860
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   10200
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line5 
      X1              =   270
      X2              =   4950
      Y1              =   3945
      Y2              =   3945
   End
   Begin VB.Line Line4 
      X1              =   5310
      X2              =   10110
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line3 
      X1              =   5400
      X2              =   10200
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   10200
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblStartHP 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   1335
      TabIndex        =   48
      Top             =   4005
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Respawn ?"
      Height          =   375
      Left            =   5415
      TabIndex        =   43
      Top             =   1065
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "Spawn Rate (in seconds)"
      Height          =   375
      Left            =   5400
      TabIndex        =   41
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "Say"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   510
      Width           =   735
   End
   Begin VB.Label lblExpGiven 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3975
      TabIndex        =   38
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Exp Given"
      Height          =   375
      Left            =   2655
      TabIndex        =   37
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Start HP"
      Height          =   375
      Left            =   255
      TabIndex        =   36
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   9690
      TabIndex        =   32
      Top             =   3435
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Value"
      Height          =   375
      Left            =   5370
      TabIndex        =   30
      Top             =   3435
      Width           =   735
   End
   Begin VB.Label lblItemName 
      Height          =   375
      Left            =   6090
      TabIndex        =   29
      Top             =   2595
      Width           =   4095
   End
   Begin VB.Label Label11 
      Caption         =   "Item"
      Height          =   375
      Left            =   5370
      TabIndex        =   28
      Top             =   2595
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Num"
      Height          =   375
      Left            =   5370
      TabIndex        =   27
      Top             =   2955
      Width           =   855
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   9690
      TabIndex        =   26
      Top             =   2955
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Drop Item Chance 1 out of"
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "MAGI"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   2895
      Width           =   855
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   2895
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "SPD"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2505
      Width           =   855
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   2505
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "DEF"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2115
      Width           =   855
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   2115
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "STR"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   1725
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Range"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1335
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Behavior"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3315
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   75
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Sprite"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   945
      Width           =   855
   End
   Begin VB.Label lblSprite 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   945
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
    lblSprite.Caption = str(scrlSprite.value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = str(scrlRange.value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = str(scrlSTR.value)

End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = str(scrlDEF.value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = str(scrlSPEED.value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = str(scrlMAGI.value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = str(scrlNum.value)
    If scrlNum.value > 0 Then
        lblItemName.Caption = Trim(Item(scrlNum.value).Name)
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = str(scrlValue.value)
End Sub

Private Sub cmdOK_Click()
    Call NpcEditorOk
    
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

Private Sub txtExpGiven_Change()
lblExpGiven.Caption = txtExpGiven
End Sub

Private Sub txtStartHP_Change()
lblStartHP.Caption = txtStartHP
End Sub
