VERSION 5.00
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9945
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   4200
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlSpeed 
      Height          =   255
      Left            =   1200
      Max             =   30000
      Min             =   1
      TabIndex        =   38
      Top             =   3720
      Value           =   1000
      Width           =   2775
   End
   Begin VB.CheckBox chkAOE 
      Caption         =   "AOE"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   2055
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   255
      Left            =   840
      Max             =   100
      TabIndex        =   34
      Top             =   3360
      Width           =   3495
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   255
      Left            =   840
      Max             =   255
      Min             =   1
      TabIndex        =   29
      Top             =   3000
      Value           =   1
      Width           =   3495
   End
   Begin VB.Frame fraAttackSound 
      Caption         =   "Cast Sound"
      Height          =   1935
      Left            =   5040
      TabIndex        =   23
      Top             =   120
      Width           =   4815
      Begin VB.FileListBox flSound 
         Height          =   1050
         Left            =   960
         TabIndex        =   26
         Top             =   480
         Width           =   3615
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblCurrentSound 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cast Sound: None"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.PictureBox picIcon 
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
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      Top             =   1200
      Width           =   480
   End
   Begin VB.HScrollBar scrlIcon 
      Height          =   255
      Left            =   1560
      Max             =   255
      TabIndex        =   20
      Top             =   1560
      Width           =   2775
   End
   Begin VB.HScrollBar scrlMPReq 
      Height          =   255
      Left            =   1560
      Max             =   255
      TabIndex        =   16
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      ItemData        =   "frmSpellEditor.frx":0000
      Left            =   120
      List            =   "frmSpellEditor.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1455
      Left            =   5040
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Power"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraGiveItem 
      Caption         =   "Give Item"
      Height          =   1455
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlItemValue 
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   10
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblItemValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Value"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblItemNum 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  'Right Justify
      Caption         =   "1000"
      Height          =   375
      Left            =   3915
      TabIndex        =   39
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Cool Down"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblRangee 
      Caption         =   "Range"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblAnimName 
      Caption         =   "(none)"
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Anim Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblPic 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label lblPicture 
      Caption         =   "Anim"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label lblIcon 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   21
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Icon"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblMPReq2 
      Caption         =   "MP Req."
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMPReq 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   160
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "ms"
      Height          =   375
      Left            =   4320
      TabIndex        =   40
      Top             =   3720
      Width           =   495
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub cmbType_Click()
    If cmbType.ListIndex <> Spell_Type.GiveItem Then
        fraVitals.Visible = True
        fraGiveItem.Visible = False
    Else
        fraVitals.Visible = False
        fraGiveItem.Visible = True
    End If
End Sub

Private Sub cmdClear_Click()
    EditorSpellSound = vbNullString
    flSound.ListIndex = -1
    lblCurrentSound.Caption = "Cast Sound: None"
End Sub

Private Sub cmdPlay_Click()
    If LenB(EditorSpellSound) > 0 Then
        SoundPlay EditorSpellSound & SOUND_EXT
    End If
End Sub

Private Sub flSound_Click()
Dim FileName() As String
Dim Ending As String

    If flSound.ListIndex < 0 Then Exit Sub

    FileName = Split(flSound.List(flSound.ListIndex), ".", , vbTextCompare)
    
    If UBound(FileName) > 1 Then
        MsgBox "Invalid file name! Cannot contain any periods!"
        flSound.ListIndex = -1
        Exit Sub
    End If
    
    Ending = FileName(1)
    
    If "." & Ending <> SOUND_EXT Then
        MsgBox "." & UCase$(Ending) & " files are not supported. Please select another!"
        flSound.ListIndex = -1
        Exit Sub
    End If
    
    lblCurrentSound.Caption = "Cast Sound: " & flSound.List(flSound.ListIndex)
    EditorSpellSound = FileName(0)
    SoundPlay flSound.List(flSound.ListIndex)

End Sub

Private Sub Form_Load()
    scrlPic_Change
End Sub

Private Sub scrlIcon_Change()
    lblIcon.Caption = scrlIcon.Value
    SpellEditorBltIcon
End Sub

Private Sub scrlIcon_Scroll()
    scrlIcon_Change
End Sub

Private Sub scrlItemNum_Change()
    fraGiveItem.Caption = "Give Item " & Trim$(Item(scrlItemNum.Value).Name)
    lblItemNum.Caption = CStr(scrlItemNum.Value)
End Sub

Private Sub scrlItemNum_Scroll()
    scrlItemNum_Change
End Sub

Private Sub scrlItemValue_Change()
    lblItemValue.Caption = CStr(scrlItemValue.Value)
End Sub

Private Sub scrlItemValue_Scroll()
    scrlItemValue_Change
End Sub

Private Sub scrlMPReq_Change()
    lblMPReq.Caption = CStr(scrlMPReq.Value)
End Sub

Private Sub scrlMPReq_Scroll()
    scrlMPReq_Change
End Sub

Private Sub scrlPic_Change()
    If LenB(Trim$(Animation(scrlPic.Value).Name)) < 1 Then
        lblAnimName.Caption = "(none)"
    Else
        lblAnimName.Caption = Trim$(Animation(scrlPic.Value).Name)
    End If
    lblPic.Caption = CStr(scrlPic.Value)
End Sub

Private Sub scrlPic_Scroll()
    scrlPic_Change
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.Value)
End Sub

Private Sub scrlRange_Scroll()
    scrlRange_Change
End Sub

Private Sub scrlSpeed_Change()
    lblSpeed.Caption = scrlSpeed.Value
End Sub

Private Sub scrlSpeed_Scroll()
    scrlSpeed_Change
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.Value)
End Sub

Private Sub cmdOk_Click()
    If LenB(Trim$(txtName)) = 0 Then
        MsgBox "Name required."
    Else
        SpellEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    SpellEditorCancel
End Sub

Private Sub scrlVitalMod_Scroll()
    scrlVitalMod_Change
End Sub
