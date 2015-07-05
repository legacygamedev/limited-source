VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
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
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Index"
      Height          =   375
      Left            =   1800
      TabIndex        =   27
      Top             =   3120
      Width           =   1455
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
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   4440
      Width           =   480
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
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   480
      Width           =   480
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   17
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   960
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   1320
         Max             =   2000
         TabIndex        =   6
         Top             =   480
         Value           =   50
         Width           =   2895
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "50"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Strength"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Durability"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
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
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      ItemData        =   "frmItemEditor.frx":0000
      Left            =   120
      List            =   "frmItemEditor.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   22
         Top             =   840
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblSpellName 
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Num"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSpell 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
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
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label lblPic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Pic"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
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

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
    scrlPic.Max = NumItems
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ItemEditorOk
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    frmIndex.Show
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        scrlStrength_Change
        scrlDurability_Change
    Else
        fraEquipment.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        scrlSpell_Change
    Else
        fraSpell.Visible = False
    End If
End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = CStr(scrlPic.Value)
    Call ItemEditorBltItem
End Sub

' Equipment Data
' *********************************
Private Sub scrlDurability_Change()
    lblDurability.Caption = CStr(scrlDurability.Value)
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = CStr(scrlStrength.Value)
End Sub

' Vitals Data
' *********************************
Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.Value)
End Sub

' Spell Data
' *********************************
Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = CStr(scrlSpell.Value)
End Sub
