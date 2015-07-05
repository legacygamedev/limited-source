VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSkill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Skill"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Set Skill"
      TabPicture(0)   =   "frmSkill.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbSkill"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbSheet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmsheet"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame frmsheet 
         Caption         =   "Itemsheet"
         Height          =   2775
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1480
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   800
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbSheet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSkill.frx":001C
         Left            =   240
         List            =   "frmSkill.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbSkill 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSkill.frx":0020
         Left            =   240
         List            =   "frmSkill.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2175
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
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   4560
         Width           =   855
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
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Select item sheet to use :"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Skill:"
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
         TabIndex        =   5
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmskill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSheet_Change()
frmsheet.Caption = "Skillsheet " & Val(cmbSheet.ListIndex - 1)
Label1(0).Caption = item(skill(cmbSkill.ListIndex).ItemGive1num(cmbSheet.ListIndex - 1)).Name
Me.Caption = cmbSkill.ListIndex & "," & cbmsheet.ListIndex

End Sub

Private Sub cmbSheet_Click()
Dim i As Long
i = 0

'Reset the labels
Label1(0).Caption = ""
Label1(1).Caption = ""
Label1(2).Caption = ""
Label1(3).Caption = ""
Label1(4).Caption = ""

If cmbSheet.ListIndex > 0 Then
    frmsheet.Caption = "Skillsheet " & Val(cmbSheet.ListIndex)
    
    If 0 + skill(cmbSkill.ListIndex + 1).itemequiped(cmbSheet.ListIndex) <> 0 Then
        Label1(i).Caption = "Item equiped: " & item(skill(cmbSkill.ListIndex + 1).itemequiped(cmbSheet.ListIndex)).Name
        i = i + 1
    End If
    
    If skill(cmbSkill.ListIndex + 1).ItemGive1num(cmbSheet.ListIndex) <> 0 Then
        Label1(i).Caption = "First item given: " & item(skill(cmbSkill.ListIndex + 1).ItemGive1num(cmbSheet.ListIndex)).Name
        i = i + 1
    End If
    
    If skill(cmbSkill.ListIndex + 1).ItemGive2num(cmbSheet.ListIndex) <> 0 Then
        Label1(i).Caption = "Second item given: " & item(skill(cmbSkill.ListIndex + 1).ItemGive2num(cmbSheet.ListIndex)).Name
        i = i + 1
    End If
    
    If skill(cmbSkill.ListIndex + 1).ItemTake1num(cmbSheet.ListIndex) <> 0 Then
        Label1(i).Caption = "First item taken: " & item(skill(cmbSkill.ListIndex + 1).ItemTake1num(cmbSheet.ListIndex)).Name
        i = i + 1
    End If
    
    If skill(cmbSkill.ListIndex + 1).ItemTake2num(cmbSheet.ListIndex) <> 0 Then
        Label1(i).Caption = "Second item taken: " & item(skill(cmbSkill.ListIndex + 1).ItemTake2num(cmbSheet.ListIndex)).Name
        i = i + 1
    End If
Else
    frmsheet.Caption = "All sheets"
End If

End Sub

Private Sub cmdCancel_Click()
Me.Visible = False
frmskill.cmbSkill.Clear
frmskill.cmbSheet.Clear
End Sub

Private Sub cmdOk_Click()
skill1 = cmbSkill.ListIndex + 1
skill2 = cmbSheet.ListIndex
Me.Visible = False
frmskill.cmbSkill.Clear
frmskill.cmbSheet.Clear
End Sub

