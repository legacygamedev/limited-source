VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   9270
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   5895
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpellEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   618
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1587
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Spell"
      TabPicture(0)   =   "frmSpellEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblRange"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSpellAnim"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSpellTime"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSpellDone"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkArea"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "info"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "scrlSound"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "scrlVitalMod"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbType"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdCancel"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdOk"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "scrlRange"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "scrlSpellAnim"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "scrlSpellTime"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "scrlSpellDone"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "picSpell"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   3720
         TabIndex        =   33
         Top             =   6600
         Width           =   1095
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   6600
         Width           =   480
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   270
         Left            =   240
         Max             =   10
         Min             =   1
         TabIndex        =   31
         Top             =   8040
         Value           =   1
         Width           =   5175
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   270
         Left            =   240
         Max             =   500
         Min             =   40
         TabIndex        =   30
         Top             =   7440
         Value           =   40
         Width           =   5175
      End
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   270
         Left            =   240
         Max             =   2000
         TabIndex        =   29
         Top             =   6840
         Value           =   1
         Width           =   4575
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   270
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   25
         Top             =   6120
         Value           =   1
         Width           =   5175
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
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   8520
         Width           =   1215
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
         Height          =   300
         Left            =   4200
         TabIndex        =   21
         Top             =   8520
         Width           =   1230
      End
      Begin VB.ComboBox cmbType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmSpellEditor.frx":0FDE
         Left            =   240
         List            =   "frmSpellEditor.frx":0FF4
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4200
         Width           =   5175
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   270
         Left            =   240
         Max             =   1000
         TabIndex        =   14
         Top             =   4920
         Width           =   5175
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   270
         Left            =   240
         Max             =   100
         TabIndex        =   13
         Top             =   5520
         Width           =   5175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Spell Qualities"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   5175
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   500
            TabIndex        =   8
            Top             =   600
            Value           =   1
            Width           =   4935
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   270
            Left            =   120
            Max             =   1000
            TabIndex        =   7
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level Required:"
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
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP Cost:"
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
            Left            =   180
            TabIndex        =   11
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Spell"
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
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   1050
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
            Left            =   840
            TabIndex        =   9
            Top             =   960
            Width           =   75
         End
      End
      Begin VB.Frame info 
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox cmbClassReq 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":1059
            Left            =   120
            List            =   "frmSpellEditor.frx":105B
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   4905
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   4875
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Required"
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
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Area Effect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   34
         Top             =   8520
         Width           =   1215
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Animation 1 Time"
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
         Left            =   240
         TabIndex        =   28
         Top             =   7800
         Width           =   1515
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time: 40"
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
         Left            =   240
         TabIndex        =   27
         Top             =   7200
         Width           =   555
      End
      Begin VB.Label lblSpellAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: 0"
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
         Left            =   240
         TabIndex        =   26
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label lblRange 
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
         Left            =   720
         TabIndex        =   24
         Top             =   5880
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Range:"
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
         Left            =   240
         TabIndex        =   23
         Top             =   5880
         Width           =   780
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Spell Type"
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
         Left            =   240
         TabIndex        =   20
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblVitalMod 
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
         TabIndex        =   19
         Top             =   4680
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vital Mod:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   4680
         Width           =   660
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   5280
         Width           =   780
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Sound"
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
         Left            =   840
         TabIndex        =   16
         Top             =   5280
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Done As Long
Private Time As Long
Private SpellVar As Long



Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = STR(scrlCost.Value)
End Sub

Private Sub scrlLevelReq_Change()
    If STR(scrlLevelReq.Value) = 0 Then
        lblLevelReq.Caption = "God's Only Spell"
    Else
        lblLevelReq.Caption = STR(scrlLevelReq.Value)
    End If
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSound_Change()
If STR(scrlSound.Value) = 0 Then
    lblSound.Caption = "No Sound"
Else
    lblSound.Caption = STR(scrlSound.Value)
    Call PlaySound("magic" & scrlSound.Value & ".wav")
End If
End Sub

Private Sub scrlSpellAnim_Change()
    lblSpellAnim.Caption = "Anim: " & scrlSpellAnim.Value
    Done = 0
End Sub

Private Sub scrlSpellDone_Change()
Dim String2 As String
    String2 = "Times"
    If scrlSpellDone.Value = 1 Then String2 = "Time"
    lblSpellDone.Caption = "Cycle Animation " & scrlSpellDone.Value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Time: " & scrlSpellTime.Value
    Done = 0
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub Timer1_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim SpellDone As Long
Dim SpellAnim As Long
Dim SpellTime As Long

SpellDone = scrlSpellDone.Value
SpellAnim = scrlSpellAnim.Value
SpellTime = scrlSpellTime.Value

If SpellAnim <= 0 Then Exit Sub
If Done = SpellDone Then Exit Sub

    With dRECT
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With

    If SpellVar > 10 Then
        Done = Done + 1
        SpellVar = 0
    End If
    If GetTickCount > Time + SpellTime Then
        Time = GetTickCount
        SpellVar = SpellVar + 1
    End If

    If DD_SpellAnim Is Nothing Then
    Else
        With sRECT
            .Top = SpellAnim * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = SpellVar * PIC_X
            .Right = .Left + PIC_X
        End With
        
        Call DD_SpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
        picSpell.Refresh
    End If
End Sub
