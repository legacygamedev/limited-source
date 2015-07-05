VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmQuestEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest Editor"
   ClientHeight    =   8400
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraExpTypes 
      Caption         =   "Experience Rewards"
      Height          =   2535
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.HScrollBar scrlQuestExpReward 
         Height          =   270
         Left            =   1800
         Max             =   32000
         TabIndex        =   54
         Top             =   480
         Value           =   1
         Width           =   2055
      End
      Begin VB.CheckBox chkBows 
         Caption         =   "Bows"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkAxes 
         Caption         =   "Axes Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox chkPoles 
         Caption         =   "Poles Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkBW 
         Caption         =   "B. Weapons Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkSB 
         Caption         =   "Small Blades Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkLB 
         Caption         =   "Large Blades Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Experience To Give:"
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
         TabIndex        =   56
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label lbExpReward 
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
         Left            =   3240
         TabIndex        =   55
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7920
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   2895
      Left            =   0
      TabIndex        =   21
      Top             =   4920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Give On Start"
      TabPicture(0)   =   "frmQuestEditor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "scrlstartval"
      Tab(0).Control(1)=   "scrlstartnum"
      Tab(0).Control(2)=   "chkstart"
      Tab(0).Control(3)=   "lblstartval"
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(6)=   "lblstartitem"
      Tab(0).Control(7)=   "Label8"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Quest Item"
      TabPicture(1)   =   "frmQuestEditor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "scrlquestitem"
      Tab(1).Control(1)=   "scrlquestvalue"
      Tab(1).Control(2)=   "lblquestval"
      Tab(1).Control(3)=   "lblquestitem"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label11"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Reward Item"
      TabPicture(2)   =   "frmQuestEditor.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblrewval"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblrewitem"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label18"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "scrlrewitem"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "scrlrewval"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.HScrollBar scrlrewval 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   41
         Top             =   2040
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlrewitem 
         Height          =   255
         Left            =   120
         Max             =   500
         Min             =   1
         TabIndex        =   40
         Top             =   1320
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlquestitem 
         Height          =   255
         Left            =   -74880
         Max             =   500
         Min             =   1
         TabIndex        =   34
         Top             =   1560
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlquestvalue 
         Height          =   255
         Left            =   -74880
         Min             =   1
         TabIndex        =   33
         Top             =   2280
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlstartval 
         Height          =   255
         Left            =   -74880
         Min             =   1
         TabIndex        =   28
         Top             =   2520
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlstartnum 
         Height          =   255
         Left            =   -74880
         Max             =   500
         Min             =   1
         TabIndex        =   25
         Top             =   1800
         Value           =   1
         Width           =   3975
      End
      Begin VB.CheckBox chkstart 
         Caption         =   "Give Item On Start"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label18 
         Caption         =   "Value :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblrewitem 
         Caption         =   "lblrewitem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblrewval 
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label14 
         Caption         =   "This is the item the user will receive once he has succesfully beaten the quest!"
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblquestval 
         Height          =   255
         Left            =   -74280
         TabIndex        =   38
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblquestitem 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label13 
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label12 
         Caption         =   "Value :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "This is the item the user must give to the NPC to succesfully complete the quest!"
         Height          =   975
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblstartval 
         Height          =   255
         Left            =   -74280
         TabIndex        =   29
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Value :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblstartitem 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label8 
         Caption         =   "If you want to give the user an item when he starts the quest, here's the place to do it [Example : Give the user a special key] :"
         Height          =   615
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   3975
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Before/After"
      TabPicture(0)   =   "frmQuestEditor.frx":0054
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtbefore"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtafter"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Timer1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Start/End "
      TabPicture(1)   =   "frmQuestEditor.frx":0070
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "txtstart"
      Tab(1).Control(3)=   "txtend"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "During"
      TabPicture(2)   =   "frmQuestEditor.frx":008C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtnotitem"
      Tab(2).Control(1)=   "txtduring"
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(3)=   "Label6"
      Tab(2).ControlCount=   4
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2160
         Top             =   1920
      End
      Begin VB.TextBox txtnotitem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox txtduring 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   18
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtend 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   16
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtstart 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtafter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtbefore 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "This is what the NPC will say if the user does NOT have the item [Example : Please come back when you have gotten the carrot!] :"
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "This is what the NPC will ask the user when he asks the user if he has the item [Example : Do you have the carrot?] :"
         Height          =   975
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   $"frmQuestEditor.frx":00A8
         Height          =   735
         Left            =   -74880
         TabIndex        =   15
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   $"frmQuestEditor.frx":013E
         Height          =   855
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   $"frmQuestEditor.frx":01DE
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "NPC Says this before player meets requirements and before player accepts quest. In other words, what the NPC normally says :"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmQuestEditor.frx":0284
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLevel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstclass"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkcls"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrllvl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chklvl"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExpTypes"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmdExpTypes 
         Caption         =   "Exp Rewards"
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chklvl 
         Caption         =   "Quest Level Requirement"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.HScrollBar scrllvl 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         Max             =   500
         Min             =   1
         TabIndex        =   5
         Top             =   720
         Value           =   1
         Width           =   1095
      End
      Begin VB.CheckBox chkcls 
         Caption         =   "Quest Class Requirement"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ListBox lstclass 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "500"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Quest Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmQuestEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check33_Click()

End Sub

Private Sub Check36_Click()

End Sub

Private Sub chkcls_Click()
If frmQuestEditor.chkcls.Value = 1 Then
frmQuestEditor.lstclass.Enabled = True
Else
frmQuestEditor.lstclass.Enabled = False
End If
End Sub

Private Sub chklvl_Click()
If frmQuestEditor.chklvl.Value = 1 Then
frmQuestEditor.scrllvl.Enabled = True
Else
frmQuestEditor.scrllvl.Enabled = False
End If
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub chkstart_Click()
If chkstart.Value = 1 Then
        frmQuestEditor.scrlstartnum.Enabled = True
        frmQuestEditor.scrlstartval.Enabled = True
        frmQuestEditor.lblstartitem = frmQuestEditor.scrlstartnum.Value & ":" & Item(frmQuestEditor.scrlstartnum.Value).Name
        frmQuestEditor.lblstartval = Quest(EditorIndex).Startval
    Else
        frmQuestEditor.scrlstartnum.Value = 1
        frmQuestEditor.scrlstartval.Value = 1
        frmQuestEditor.scrlstartnum.Enabled = False
        frmQuestEditor.scrlstartval.Enabled = False
        frmQuestEditor.lblstartitem = "Disabled"
        frmQuestEditor.lblstartval = "Disabled"
End If
End Sub

Private Sub cmdExpTypes_Click()
If fraExpTypes.Visible = False Then
frmQuestEditor.fraExpTypes.Visible = True
Else
frmQuestEditor.fraExpTypes.Visible = False
End If
End Sub

Private Sub Command1_Click()
Call QuestEditorOk
End Sub

Private Sub Command2_Click()
Call QuestEditorCancel
End Sub

Private Sub Command3_Click()
If fraExpTypes.Visible = False Then
frmQuestEditor.fraExpTypes.Visible = True
Else
frmQuestEditor.fraExpTypes.Visible = False
End If
End Sub

Private Sub form_load()
frmQuestEditor.scrlstartnum.Max = MAX_ITEMS
frmQuestEditor.lblrewitem.Caption = scrlrewitem.Value & ":" & Item(scrlrewitem.Value).Name
frmQuestEditor.lblquestitem.Caption = frmQuestEditor.scrlquestitem.Value & ":" & Item(scrlquestitem.Value).Name
End Sub

Private Sub scrllvl_Change()
frmQuestEditor.lblLevel.Caption = scrllvl.Value
End Sub

Private Sub scrlQuestExpReward_Change()
frmQuestEditor.lbExpReward.Caption = frmQuestEditor.scrlQuestExpReward.Value
End Sub

Private Sub scrlquestitem_Change()
frmQuestEditor.lblquestitem.Caption = frmQuestEditor.scrlquestitem.Value & ":" & Item(scrlquestitem.Value).Name
End Sub

Private Sub scrlquestvalue_Change()
frmQuestEditor.lblquestval.Caption = frmQuestEditor.scrlquestvalue.Value
End Sub

Private Sub scrlrewitem_Change()
frmQuestEditor.lblrewitem.Caption = scrlrewitem.Value & ":" & Item(scrlrewitem.Value).Name
End Sub

Private Sub scrlrewval_Change()
frmQuestEditor.lblrewval.Caption = frmQuestEditor.scrlrewval.Value
End Sub


Private Sub scrlstartnum_Change()
frmQuestEditor.lblstartitem.Caption = scrlstartnum.Value & ":" & Item(scrlstartnum.Value).Name
End Sub

Private Sub scrlstartval_Change()
frmQuestEditor.lblstartval.Caption = frmQuestEditor.scrlstartval.Value
End Sub

Private Sub Timer1_Timer()
If frmQuestEditor.chkstart.Value = 1 Then
If Item(frmQuestEditor.scrlstartnum.Value).Type = 12 Then
frmQuestEditor.scrlstartval.Enabled = True
frmQuestEditor.lblstartval.Caption = frmQuestEditor.scrlstartval.Value
Else
frmQuestEditor.scrlstartval.Enabled = False
frmQuestEditor.lblstartval.Caption = "1"
End If
End If
If Item(frmQuestEditor.scrlquestitem.Value).Type = 12 Then
frmQuestEditor.scrlquestvalue.Enabled = True
frmQuestEditor.lblquestval.Caption = frmQuestEditor.scrlquestvalue.Value
Else
frmQuestEditor.scrlquestvalue.Enabled = False
frmQuestEditor.lblquestval.Caption = "1"
End If
If Item(frmQuestEditor.scrlrewitem.Value).Type = 12 Then
frmQuestEditor.scrlrewval.Enabled = True
frmQuestEditor.lblrewval.Caption = frmQuestEditor.scrlrewval.Value
Else
frmQuestEditor.scrlrewval.Enabled = False
frmQuestEditor.lblrewval.Caption = "1"
End If
End Sub
