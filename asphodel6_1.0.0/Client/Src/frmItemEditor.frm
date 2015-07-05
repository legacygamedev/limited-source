VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
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
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlAnim 
      Height          =   255
      Left            =   5640
      Max             =   255
      TabIndex        =   72
      Top             =   960
      Value           =   1
      Width           =   3255
   End
   Begin VB.PictureBox picPic 
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
      Left            =   8880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   240
      Width           =   480
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   255
      Left            =   5400
      Max             =   255
      TabIndex        =   17
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      ItemData        =   "frmItemEditor.frx":0000
      Left            =   120
      List            =   "frmItemEditor.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1575
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.HScrollBar scrlVitalMod3 
         Height          =   255
         Left            =   1080
         Min             =   -32767
         TabIndex        =   87
         Top             =   1080
         Width           =   2415
      End
      Begin VB.HScrollBar scrlVitalMod2 
         Height          =   255
         Left            =   1080
         Min             =   -32767
         TabIndex        =   84
         Top             =   720
         Width           =   2415
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   1080
         Min             =   -32767
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblVitalMod3 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   300
         Left            =   3480
         TabIndex        =   88
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "SP Mod"
         Height          =   300
         Left            =   120
         TabIndex        =   86
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblVitalMod2 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   3480
         TabIndex        =   85
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "MP Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "HP Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1455
      Left            =   5040
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   840
         Max             =   255
         Min             =   1
         TabIndex        =   21
         Top             =   840
         Value           =   1
         Width           =   3015
      End
      Begin VB.Label lblSpellName 
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Num.:"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame FraRequirements 
      Caption         =   "Requirements"
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   48
      Top             =   2040
      Width           =   4695
      Begin VB.HScrollBar scrlAccess 
         Height          =   255
         Left            =   960
         Max             =   4
         TabIndex        =   68
         Top             =   960
         Width           =   3015
      End
      Begin VB.HScrollBar scrlRequires 
         Height          =   255
         Index           =   3
         Left            =   1320
         Max             =   1000
         TabIndex        =   61
         Top             =   2040
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRequires 
         Height          =   255
         Index           =   2
         Left            =   1320
         Max             =   1000
         TabIndex        =   60
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRequires 
         Height          =   255
         Index           =   1
         Left            =   1320
         Max             =   1000
         TabIndex        =   59
         Top             =   1560
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRequires 
         Height          =   255
         Index           =   0
         Left            =   1320
         Max             =   1000
         TabIndex        =   58
         Top             =   1320
         Width           =   2655
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   52
         Top             =   720
         Value           =   1
         Width           =   3015
      End
      Begin VB.ComboBox cmbClass 
         Height          =   360
         ItemData        =   "frmItemEditor.frx":0066
         Left            =   840
         List            =   "frmItemEditor.frx":006D
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label21 
         Caption         =   "Access:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblAccess 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   66
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblRequires 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   65
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblRequires 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   64
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblRequires 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   63
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblRequires 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   62
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Magic:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Defense:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Strength:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   240
         Left            =   4350
         TabIndex        =   53
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label16 
         Caption         =   "Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   400
         Width           =   1215
      End
   End
   Begin VB.Frame fraWorth 
      Caption         =   "Worth"
      Height          =   1215
      Left            =   120
      TabIndex        =   74
      Top             =   840
      Width           =   4695
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   1080
         TabIndex        =   81
         Top             =   720
         Width           =   3015
      End
      Begin VB.HScrollBar scrlWorthItem 
         Height          =   255
         Left            =   840
         Max             =   0
         TabIndex        =   75
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4080
         TabIndex        =   82
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Amount:"
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblItemName 
         Caption         =   "(none)"
         Height          =   255
         Left            =   1440
         TabIndex        =   78
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label25 
         Caption         =   "Item:"
         Height          =   375
         Left            =   240
         TabIndex        =   77
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWorthItem 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4080
         TabIndex        =   76
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3615
      Left            =   5040
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame FraBonuses 
         Caption         =   "Bonuses"
         Height          =   2055
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   4095
         Begin VB.HScrollBar scrlBonusStat 
            Height          =   255
            Index           =   4
            Left            =   1080
            Max             =   1000
            TabIndex        =   30
            Top             =   1680
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusStat 
            Height          =   255
            Index           =   3
            Left            =   1080
            Max             =   1000
            TabIndex        =   32
            Top             =   1440
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusStat 
            Height          =   255
            Index           =   2
            Left            =   1080
            Max             =   1000
            TabIndex        =   31
            Top             =   1200
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusStat 
            Height          =   255
            Index           =   1
            Left            =   1080
            Max             =   1000
            TabIndex        =   28
            Top             =   960
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusVital 
            Height          =   255
            Index           =   3
            Left            =   1080
            Max             =   1000
            TabIndex        =   39
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusVital 
            Height          =   255
            Index           =   2
            Left            =   1080
            Max             =   1000
            TabIndex        =   41
            Top             =   480
            Width           =   2175
         End
         Begin VB.HScrollBar scrlBonusVital 
            Height          =   255
            Index           =   1
            Left            =   1080
            Max             =   1000
            TabIndex        =   40
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblBonusVital 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   47
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblBonusVital 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblBonusVital 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   45
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "SP:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "MP:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "HP:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Magic:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Defense:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblBonusStat 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   35
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblBonusStat 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   34
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblBonusStat 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   3240
            TabIndex        =   33
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblBonusStat 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   29
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Strength:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Value           =   1
         Width           =   2655
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   1080
         Max             =   1000
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "End."
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Power"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         UseMnemonic     =   0   'False
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
   Begin VB.Label lblAnim 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   8880
      TabIndex        =   73
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "Anim:"
      Height          =   375
      Left            =   5040
      TabIndex        =   71
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAnimName 
      Caption         =   "(none)"
      Height          =   255
      Left            =   6240
      TabIndex        =   70
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label22 
      Caption         =   "Anim Name:"
      Height          =   255
      Left            =   5040
      TabIndex        =   69
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblPic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Pic"
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub cmdOk_Click()

    If LenB(Trim$(txtName)) = 0 Then
        MsgBox "Name required.", vbOKOnly + vbCritical, Game_Name
        Exit Sub
    End If
    
    If scrlWorthItem.Value > 0 Then
        If Item(scrlWorthItem.Value).Type <> ItemType.Currency_ Then
            MsgBox "The item worth must be a currency type item!", vbOKOnly + vbCritical, Game_Name
            Exit Sub
        End If
    End If
    
    ItemEditorOk
    
End Sub

Private Sub cmdCancel_Click()
    ItemEditorCancel
End Sub

Private Sub cmbType_Click()

    fraEquipment.Visible = ((cmbType.ListIndex >= ItemType.Weapon_) And (cmbType.ListIndex <= ItemType.Shield_))
    fraVitals.Visible = (cmbType.ListIndex = ItemType.Potion)
    fraSpell.Visible = (cmbType.ListIndex = ItemType.Spell_)
    
    If cmbType.ListIndex = ItemType.Currency_ Or cmbType.ListIndex = 0 Then Exit Sub
    
    fraWorth.Visible = True
    
End Sub

Private Sub Form_Load()
    scrlAnim_Change
    scrlSpell_Change
    scrlWorthItem.Max = MAX_ITEMS
End Sub

Private Sub scrlAccess_Change()
    lblAccess.Caption = scrlAccess.Value
End Sub

Private Sub scrlAccess_Scroll()
    scrlAccess_Change
End Sub

Private Sub scrlAmount_Change()
    lblAmount.Caption = scrlAmount.Value
End Sub

Private Sub scrlAmount_Scroll()
    scrlAmount_Change
End Sub

Private Sub scrlAnim_Change()
On Error Resume Next

    If LenB(Trim$(Animation(scrlAnim.Value).Name)) < 1 Then
        lblAnimName.Caption = "(none)"
    Else
        lblAnimName.Caption = Trim$(Animation(scrlAnim.Value).Name)
    End If
    
    lblAnim.Caption = scrlAnim.Value
End Sub

Private Sub scrlAnim_Scroll()
    scrlAnim_Change
End Sub

Private Sub scrlBonusStat_Change(Index As Integer)
    lblBonusStat(Index).Caption = scrlBonusStat(Index).Value
End Sub

Private Sub scrlBonusStat_Scroll(Index As Integer)
    scrlBonusStat_Change Index
End Sub

Private Sub scrlBonusVital_Change(Index As Integer)
    lblBonusVital(Index).Caption = scrlBonusVital(Index).Value
End Sub

Private Sub scrlBonusVital_Scroll(Index As Integer)
    scrlBonusVital_Change Index
End Sub

Private Sub scrlDurability_Scroll()
    scrlDurability_Change
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = scrlLevel.Value
End Sub

Private Sub scrlLevel_Scroll()
    scrlLevel_Change
End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = CStr(scrlPic.Value)
    ItemEditorBltItem
End Sub

' Equipment Data
' *********************************
Private Sub scrlDurability_Change()
    If scrlDurability.Value = 0 Then lblDurability.Caption = "End.": Exit Sub
    lblDurability.Caption = CStr(scrlDurability.Value)
End Sub

Private Sub scrlPic_Scroll()
    scrlPic_Change
End Sub

Private Sub scrlRequires_Change(Index As Integer)
    lblRequires(Index).Caption = scrlRequires(Index).Value
End Sub

Private Sub scrlRequires_Scroll(Index As Integer)
    scrlRequires_Change Index
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = CStr(scrlStrength.Value)
End Sub

Private Sub scrlStrength_Scroll()
    scrlStrength_Change
End Sub

' Vitals Data
' *********************************
Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.Value)
End Sub

Private Sub scrlVitalMod2_Change()
    lblVitalMod2.Caption = scrlVitalMod2.Value
End Sub

Private Sub scrlVitalMod3_Change()
    lblVitalMod3.Caption = scrlVitalMod3.Value
End Sub

' Spell Data
' *********************************
Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    If LenB(lblSpellName.Caption) < 1 Then lblSpellName.Caption = "(empty)"
    lblSpell.Caption = CStr(scrlSpell.Value)
End Sub

Private Sub scrlSpell_Scroll()
    scrlSpell_Change
End Sub

Private Sub scrlWorthItem_Change()
On Error Resume Next

    lblWorthItem.Caption = scrlWorthItem.Value
    
    If scrlWorthItem.Value = 0 Then
        lblItemName.Caption = "(none)"
        Exit Sub
    End If
    
    If LenB(Trim$(Item(scrlWorthItem.Value).Name)) > 0 Then
        lblItemName.Caption = Trim$(Item(scrlWorthItem.Value).Name)
    Else
        lblItemName.Caption = "(none)"
    End If
    
End Sub

Private Sub scrlWorthItem_Scroll()
    scrlWorthItem_Change
End Sub
