VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   0  'None
   Caption         =   "Item Editor"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraArrows 
      Caption         =   "Arrows"
      Height          =   1815
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   4335
      Begin VB.HScrollBar scrlArrowQuantity 
         Height          =   255
         Left            =   1080
         Max             =   1000
         TabIndex        =   28
         Top             =   360
         Width           =   2535
      End
      Begin VB.HScrollBar scrlArrowRange 
         Height          =   255
         Left            =   1080
         Max             =   30
         TabIndex        =   27
         Top             =   720
         Width           =   2535
      End
      Begin VB.HScrollBar scrlArrowAnim 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.PictureBox picArrowAnim 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3360
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   25
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label13 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Range"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Animation"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblArrowQuantity 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3720
         TabIndex        =   31
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblArrowRange 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblArrowAnim 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   1200
         Width           =   90
      End
   End
   Begin VB.Frame fraSkill 
      Caption         =   "Skill"
      Height          =   1815
      Left            =   240
      TabIndex        =   35
      Top             =   2040
      Width           =   4335
      Begin VB.HScrollBar scrlSkill 
         Height          =   255
         Left            =   1080
         Max             =   50
         Min             =   1
         TabIndex        =   36
         Top             =   960
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Number"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblSkillName 
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblSkill 
         Caption         =   "1"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Frame fraCharm 
      Caption         =   "Charm"
      Height          =   1815
      Left            =   240
      TabIndex        =   41
      Top             =   2040
      Width           =   4335
      Begin VB.ComboBox cmbCharmType 
         Height          =   315
         ItemData        =   "frmItemEditor.frx":0000
         Left            =   840
         List            =   "frmItemEditor.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   480
         Width           =   3255
      End
      Begin VB.HScrollBar scrlCharmMod 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   42
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Type"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Mod"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblCharmMod 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3840
         TabIndex        =   44
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell"
      Height          =   1815
      Left            =   240
      TabIndex        =   47
      Top             =   2040
      Width           =   4335
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   50
         Min             =   1
         TabIndex        =   48
         Top             =   1080
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Number"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblSpellName 
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblSpell 
         Caption         =   "1"
         Height          =   255
         Left            =   3600
         TabIndex        =   49
         Top             =   1080
         Width           =   135
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1815
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   4335
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   1200
         Max             =   1000
         TabIndex        =   22
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblVitalMod 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3840
         TabIndex        =   23
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label6 
         Caption         =   "Vital Mod"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   1815
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   1200
         Max             =   1000
         TabIndex        =   17
         Top             =   1320
         Width           =   2535
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   1200
         Max             =   1000
         TabIndex        =   16
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbItemSubType 
         Height          =   315
         ItemData        =   "frmItemEditor.frx":00B6
         Left            =   1200
         List            =   "frmItemEditor.frx":00D2
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblStrength 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblDurability 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3840
         TabIndex        =   18
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label5 
         Caption         =   "Stat Mod"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Durability"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "SubType"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdEditItemCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdEditItemOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Timer tmrItemPic 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   4200
   End
   Begin VB.PictureBox picArrowsBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picItemsBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbItemType 
      Height          =   315
      ItemData        =   "frmItemEditor.frx":0108
      Left            =   360
      List            =   "frmItemEditor.frx":0145
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   4215
   End
   Begin VB.PictureBox picItemPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   720
      Width           =   480
   End
   Begin VB.HScrollBar scrlItemPic 
      Height          =   255
      Left            =   1080
      Max             =   50
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtItemName 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblItemPic 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   "Pic"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Private Sub cmdEditItemOK_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdEditItemCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbItemType_Click()
    If cmbItemType.ListIndex = ITEM_TYPE_TOOL Then
        fraEquipment.Visible = True
        cmbItemSubType.ListIndex = 0
        cmbItemSubType.Enabled = True
    Else
        fraEquipment.Visible = False
        cmbItemSubType.ListIndex = 0
        cmbItemSubType.Enabled = False
    End If
    
    If Not cmbItemType.ListIndex = ITEM_TYPE_TOOL Then
        If cmbItemType.ListIndex = ITEM_TYPE_WEAPON Then
            fraEquipment.Visible = True
            cmbItemSubType.ListIndex = 0
            cmbItemSubType.Enabled = True
        ElseIf cmbItemType.ListIndex >= ITEM_TYPE_ARMOR And cmbItemType.ListIndex <= ITEM_TYPE_SHIELD Then
            fraEquipment.Visible = True
            cmbItemSubType.ListIndex = 0
            cmbItemSubType.Enabled = False
        Else
            fraEquipment.Visible = False
            cmbItemSubType.ListIndex = 0
            cmbItemSubType.Enabled = False
        End If
    End If
    
    If cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP And cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP Then
        fraVitals.Visible = True
    Else
        fraVitals.Visible = False
    End If
    
    If cmbItemType.ListIndex >= ITEM_TYPE_AMULET And cmbItemType.ListIndex <= ITEM_TYPE_RING Then
        fraCharm.Visible = True
    Else
        fraCharm.Visible = False
    End If
    
    If cmbItemType.ListIndex = ITEM_TYPE_SPELL Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If cmbItemType.ListIndex = ITEM_TYPE_SKILL Then
        fraSkill.Visible = True
    Else
        fraSkill.Visible = False
    End If
    
    If cmbItemType.ListIndex = ITEM_TYPE_ARROW Then
        fraArrows.Visible = True
    Else
        fraArrows.Visible = False
    End If
End Sub

Private Sub scrlItemPic_Change()
    lblItemPic.Caption = STR(scrlItemPic.Value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = STR(scrlDurability.Value)
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.Value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub scrlSkill_Change()
    lblSkillName.Caption = Trim(Skill(scrlSkill.Value).Name)
    lblSkill.Caption = STR(scrlSkill.Value)
End Sub

Private Sub cmbCharmType_Click()
    If cmbCharmType.ListIndex >= CHARM_TYPE_ADDHP And cmbCharmType.ListIndex <= CHARM_TYPE_ADDSPEED Then
        frmItemEditor.Label10.Caption = "Mod"
    Else
        frmItemEditor.Label10.Caption = "Mod %"
    End If
End Sub

Private Sub scrlCharmMod_Change()
    lblCharmMod.Caption = STR(scrlCharmMod.Value)
End Sub

Private Sub scrlArrowQuantity_Change()
    lblArrowQuantity.Caption = STR(scrlArrowQuantity.Value)
End Sub

Private Sub scrlArrowRange_Change()
    lblArrowRange.Caption = STR(scrlArrowRange.Value)
End Sub

Private Sub scrlArrowAnim_Change()
    lblArrowAnim.Caption = STR(scrlArrowAnim.Value)
End Sub

Private Sub tmrItemPic_Timer()
    Call ItemEditorBltItem
End Sub
