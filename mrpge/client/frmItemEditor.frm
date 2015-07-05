VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
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
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Text            =   "frmItemEditor.frx":0000
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlSpell 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   21
         Top             =   840
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblSpellName 
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Num"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSpell 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Timer tmrPic 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picItems 
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
      Height          =   2.45745e5
      Left            =   90
      ScaleHeight     =   16383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   9480
      Width           =   480
   End
   Begin VB.PictureBox picPic 
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
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   18
      Top             =   600
      Width           =   480
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   375
      Left            =   960
      Max             =   1020
      TabIndex        =   17
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2595
      TabIndex        =   15
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   105
      TabIndex        =   14
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   12
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   5370
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   4830
      Begin VB.ComboBox cmbWeaponType 
         Height          =   390
         ItemData        =   "frmItemEditor.frx":0016
         Left            =   120
         List            =   "frmItemEditor.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   4920
         Width           =   4575
      End
      Begin VB.TextBox txtPoisonVital 
         Height          =   345
         Left            =   2640
         TabIndex        =   50
         Text            =   "0"
         Top             =   4320
         Width           =   2130
      End
      Begin VB.TextBox txtPoisonLength 
         Height          =   345
         Left            =   1770
         TabIndex        =   48
         Text            =   "0"
         Top             =   3870
         Width           =   2865
      End
      Begin VB.CheckBox chkPoison 
         Caption         =   "Does weapon Poison?"
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   3570
         Width           =   4365
      End
      Begin VB.HScrollBar scrlBaseDamage 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   41
         Top             =   3120
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlCha 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   38
         Top             =   2520
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlWiz 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   2160
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlCon 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   32
         Top             =   1800
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDex 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   29
         Top             =   1440
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlIntel 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   26
         Top             =   1080
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   720
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   6
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label14 
         Caption         =   "poison damage per tick"
         Height          =   345
         Left            =   135
         TabIndex        =   51
         Top             =   4350
         Width           =   2910
      End
      Begin VB.Label Label12 
         Caption         =   "Time poisoned"
         Height          =   405
         Left            =   90
         TabIndex        =   49
         Top             =   3885
         Width           =   1725
      End
      Begin VB.Label Label10 
         Caption         =   "Base Damage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblBaseDamage 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   42
         Top             =   3120
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label17 
         Caption         =   "Cha"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblCha 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   39
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Wiz"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblWiz 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   36
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Con"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   33
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Dex"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblDex 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Intel"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIntel 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Strength"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Durability"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox cmbType 
      Height          =   390
      ItemData        =   "frmItemEditor.frx":0034
      Left            =   120
      List            =   "frmItemEditor.frx":0068
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label lblPic 
      Caption         =   "0"
      Height          =   345
      Left            =   3825
      TabIndex        =   46
      Top             =   660
      Width           =   690
   End
   Begin VB.Label Label8 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Pic"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
    Else
        fraEquipment.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Or (cmbType.ListIndex = ITEM_TYPE_POTIONADDPP) Then
        fraVitals.Visible = True
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        fraSpell.Caption = "Spell Data"
    Else
        fraSpell.Visible = False
    End If
    If (cmbType.ListIndex = ITEM_TYPE_PRAYER) Then
        fraSpell.Visible = True
        fraSpell.Caption = "Prayer Data"
    Else
        fraSpell.Visible = False
    End If
End Sub



Private Sub Form_Load()
picItems = LoadPicture(App.Path & "\data\bmp\Items.bmp")
End Sub





Private Sub lblBaseDamage_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblBaseDamage.Caption
lblBaseDamage.Caption = InputBox("Enter Amount:")
scrlBaseDamage.value = lblBaseDamage.Caption
Exit Sub
error:
scrlBaseDamage.value = prevNum
lblBaseDamage.Caption = prevNum
End Sub

Private Sub lblCha_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblCha.Caption
lblCha.Caption = InputBox("Enter Amount:")
scrlCha.value = lblCha.Caption
Exit Sub
error:
scrlCha.value = prevNum
lblCha.Caption = prevNum
End Sub

Private Sub lblCon_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblCon.Caption
lblCon.Caption = InputBox("Enter Amount:")
scrlCon.value = lblCon.Caption
Exit Sub
error:
scrlCon.value = prevNum
lblCon.Caption = prevNum
End Sub

Private Sub lblDex_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblDex.Caption
lblDex.Caption = InputBox("Enter Amount:")
scrlDex.value = lblDex.Caption
Exit Sub
error:
scrlDex.value = prevNum
lblDex.Caption = prevNum
End Sub

Private Sub lblDurability_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblDurability.Caption
lblDurability.Caption = InputBox("Enter Amount:")
scrlDurability.value = lblDurability.Caption
Exit Sub
error:
scrlDurability.value = prevNum
lblDurability.Caption = prevNum
End Sub

Private Sub lblIntel_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblIntel.Caption
lblIntel.Caption = InputBox("Enter Amount:")
scrlIntel.value = lblIntel.Caption
Exit Sub
error:
scrlIntel.value = prevNum
lblIntel.Caption = prevNum
End Sub

Private Sub lblPic_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblPic.Caption
lblPic.Caption = InputBox("Enter Amount:")
scrlPic.value = lblPic.Caption
Exit Sub
error:
scrlPic.value = prevNum
lblPic.Caption = prevNum
End Sub

Private Sub lblSpell_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblSpell.Caption
lblSpell.Caption = InputBox("Enter Amount:")
scrlSpell.value = lblSpell.Caption
Exit Sub
error:
scrlSpell.value = prevNum
lblSpell.Caption = prevNum
End Sub

Private Sub lblStrength_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblStrength.Caption
lblStrength.Caption = InputBox("Enter Amount:")
scrlStrength.value = lblStrength.Caption
Exit Sub
error:
scrlStrength.value = prevNum
lblStrength.Caption = prevNum
End Sub

Private Sub lblVitalMod_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblVitalMod.Caption
lblVitalMod.Caption = InputBox("Enter Amount:")
scrlVitalMod.value = lblVitalMod.Caption
Exit Sub
error:
scrlVitalMod.value = prevNum
lblVitalMod.Caption = prevNum
End Sub

Private Sub lblWiz_Click()
On Error GoTo error:
Dim prevNum As Long
prevNum = lblWiz.Caption
lblWiz.Caption = InputBox("Enter Amount:")
scrlWiz.value = lblWiz.Caption
Exit Sub
error:
scrlWiz.value = prevNum
lblWiz.Caption = prevNum
End Sub

Private Sub scrlBaseDamage_Change()
    lblBaseDamage.Caption = str(scrlBaseDamage.value)
End Sub

Private Sub scrlCha_Change()
    lblCha.Caption = str(scrlCha.value)
End Sub

Private Sub scrlCon_Change()
    lblCon.Caption = str(scrlCon.value)
End Sub

Private Sub scrlDex_Change()
    lblDex.Caption = str(scrlDex.value)
End Sub

Private Sub scrlIntel_Change()
    lblIntel.Caption = str(scrlIntel.value)
End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = str(scrlPic.value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = str(scrlVitalMod.value)
End Sub

Private Sub scrlVitalAdd_Change()
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = str(scrlDurability.value)
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = str(scrlStrength.value)
End Sub

Private Sub scrlSpell_Change()
On Local Error Resume Next
If (fraSpell.Caption = "Spell Data") Then
    lblSpellName.Caption = Trim(Spell(scrlSpell.value).Name)
    lblSpell.Caption = str(scrlSpell.value)
Else
    lblSpellName.Caption = Trim(Prayer(scrlSpell.value).Name)
    lblSpell.Caption = str(scrlSpell.value)
End If
End Sub

Private Sub scrlWiz_Change()
    lblWiz.Caption = str(scrlWiz.value)
End Sub

Private Sub tmrPic_Timer()
    Call ItemEditorBltItem
End Sub

