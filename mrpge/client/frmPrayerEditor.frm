VERSION 5.00
Begin VB.Form frmPrayerEditor 
   Caption         =   "Prayer Editor"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMana 
      Height          =   390
      Left            =   960
      TabIndex        =   14
      Text            =   "10"
      Top             =   600
      Width           =   3015
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmPrayerEditor.frx":0000
      Left            =   120
      List            =   "frmPrayerEditor.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1635
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2625
      TabIndex        =   11
      Top             =   3585
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   105
      TabIndex        =   10
      Top             =   3585
      Width           =   2295
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   990
      Left            =   120
      TabIndex        =   6
      Top             =   2535
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   7
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox cmbClassReq 
      Height          =   315
      ItemData        =   "frmPrayerEditor.frx":0053
      Left            =   120
      List            =   "frmPrayerEditor.frx":0055
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
   End
   Begin VB.HScrollBar scrlLevelReq 
      Height          =   375
      Left            =   960
      Max             =   255
      Min             =   1
      TabIndex        =   0
      Top             =   2040
      Value           =   1
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Mana:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Level"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblLevelReq 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmPrayerEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmbType_Change()
    If cmbType.ListIndex = 0 Then
        fraVitals.Visible = True
    Else
        fraVitals.Visible = False
    End If
End Sub

Private Sub scrlLevelReq_Change()
    lblLevelReq.Caption = str(scrlLevelReq.value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = str(scrlVitalMod.value)
End Sub
Private Sub cmdOK_Click()
    Call PrayerEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call PrayerEditorCancel
End Sub
