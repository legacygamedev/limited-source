VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmQuestEditor 
   Caption         =   "frmQuestEditor"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   7800
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General Information"
      TabPicture(0)   =   "frmQuestEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmGiveItem"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Requirements"
      TabPicture(1)   =   "frmQuestEditor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraQuestInformation"
      Tab(1).Control(1)=   "frameClasses"
      Tab(1).Control(2)=   "frameQuest"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Needs / Rewards"
      TabPicture(2)   =   "frmQuestEditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame frmGiveItem 
         Caption         =   "GIve Item On Quest Accept"
         Height          =   975
         Left            =   120
         TabIndex        =   57
         Top             =   6600
         Width           =   5295
         Begin VB.HScrollBar scrlGiveItem 
            Height          =   255
            Left            =   2010
            Max             =   255
            TabIndex        =   59
            Top             =   240
            Value           =   1
            Width           =   2775
         End
         Begin VB.TextBox txtGiveItemValue 
            Height          =   285
            Left            =   2280
            TabIndex        =   58
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   195
            Left            =   1680
            TabIndex        =   62
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblGiveITem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4890
            TabIndex        =   61
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblGIveItemName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Name"
            Height          =   195
            Left            =   1200
            TabIndex        =   60
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quest Needs"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   5295
         Begin VB.ComboBox cmbQuestType 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "frmQuestEditor.frx":0054
            Left            =   1905
            List            =   "frmQuestEditor.frx":005E
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox cmbNeedsIndex 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1080
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.HScrollBar scrlPart 
            Height          =   255
            Left            =   1860
            Max             =   255
            Min             =   1
            TabIndex        =   21
            Top             =   360
            Value           =   1
            Width           =   2535
         End
         Begin VB.TextBox txtRequired 
            Height          =   285
            Index           =   0
            Left            =   1875
            TabIndex        =   20
            Top             =   1560
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quest Type:"
            Height          =   195
            Left            =   960
            TabIndex        =   38
            Top             =   735
            Width           =   870
         End
         Begin VB.Label lblPart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   195
            Left            =   4500
            TabIndex        =   26
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblQuestPart 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Num:"
            Height          =   195
            Left            =   935
            TabIndex        =   25
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPC/Item Index:"
            Height          =   195
            Left            =   330
            TabIndex        =   24
            Top             =   1095
            Width           =   1470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Required Num:"
            Height          =   195
            Left            =   510
            TabIndex        =   23
            Top             =   1560
            Width           =   1290
         End
      End
      Begin VB.Frame frameQuest 
         Caption         =   "Previous Quest Requirement"
         Height          =   855
         Left            =   -74880
         TabIndex        =   18
         Top             =   1440
         Width           =   5295
         Begin VB.HScrollBar scrlPrvQuest 
            Height          =   255
            Left            =   480
            Max             =   255
            TabIndex        =   41
            Top             =   480
            Value           =   1
            Width           =   3855
         End
         Begin VB.Label lblPrevQuestName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            Height          =   195
            Left            =   480
            TabIndex        =   43
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblPrevQuest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4440
            TabIndex        =   42
            Top             =   480
            Width           =   90
         End
      End
      Begin VB.Frame frameClasses 
         Caption         =   "Class Requirement"
         Height          =   855
         Left            =   -74880
         TabIndex        =   16
         Top             =   2400
         Width           =   5295
         Begin VB.CheckBox chkClass 
            Caption         =   "Class #0"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame FraQuestInformation 
         Caption         =   "Level Requirement"
         Height          =   855
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   5295
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   480
            Max             =   255
            TabIndex        =   14
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblLevel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4440
            TabIndex        =   15
            Top             =   360
            Width           =   105
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Message Information"
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   5295
         Begin VB.TextBox txtIncomplete 
            Height          =   645
            Left            =   990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox txtComplete 
            Height          =   645
            Left            =   990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox txtDenied 
            Height          =   645
            Left            =   990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   960
            Width           =   4215
         End
         Begin VB.TextBox txtAccept 
            Height          =   645
            Left            =   990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Incomplete:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   2520
            Width           =   825
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complete:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   705
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Denied:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accept:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Quest Information"
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5295
         Begin VB.CheckBox chkRepeatable 
            Caption         =   "Repeatable?"
            Height          =   255
            Left            =   3840
            TabIndex        =   55
            Top             =   2040
            Width           =   1335
         End
         Begin VB.HScrollBar scrlEndNpc 
            Height          =   255
            Left            =   1920
            Max             =   255
            TabIndex        =   50
            Top             =   1680
            Value           =   1
            Width           =   2895
         End
         Begin VB.HScrollBar scrlStartNpc 
            Height          =   255
            Left            =   1920
            Max             =   255
            TabIndex        =   47
            Top             =   1320
            Value           =   1
            Width           =   2895
         End
         Begin VB.TextBox txtDescription 
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1785
            TabIndex        =   8
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Npc:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Npc:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblEndNpc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4920
            TabIndex        =   52
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label lblEndNpcName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            Height          =   195
            Left            =   840
            TabIndex        =   51
            Top             =   1680
            Width           =   390
         End
         Begin VB.Label lblStartNpc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   4920
            TabIndex        =   49
            Top             =   1320
            Width           =   90
         End
         Begin VB.Label lblStartNpcName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            Height          =   195
            Left            =   840
            TabIndex        =   48
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Left            =   675
            TabIndex        =   11
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Left            =   1140
            TabIndex        =   10
            Top             =   360
            Width           =   570
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Quest Rewards"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   1
         Top             =   2520
         Width           =   5295
         Begin VB.TextBox txtExp 
            Height          =   285
            Left            =   1845
            TabIndex        =   45
            Top             =   2280
            Width           =   2535
         End
         Begin VB.CheckBox chkSelectionOnly 
            Caption         =   "Select Only?"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   44
            Top             =   1560
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.TextBox txtItemValue 
            Height          =   285
            Index           =   0
            Left            =   1845
            TabIndex        =   39
            Top             =   1200
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.HScrollBar scrlRewardIndex 
            Height          =   255
            Left            =   1650
            Max             =   255
            Min             =   1
            TabIndex        =   3
            Top             =   360
            Value           =   1
            Width           =   2775
         End
         Begin VB.ComboBox cmbItemIndex 
            Height          =   315
            Index           =   0
            Left            =   1670
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   840
            Visible         =   0   'False
            Width           =   2785
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Rewarded:"
            Height          =   195
            Left            =   480
            TabIndex        =   46
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   195
            Left            =   480
            TabIndex        =   40
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label lblRewardIndex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   195
            Left            =   4530
            TabIndex        =   6
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reward Num:"
            Height          =   195
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Index:"
            Height          =   195
            Left            =   495
            TabIndex        =   4
            Top             =   855
            Width           =   1035
         End
      End
      Begin VB.Label Label16 
         Caption         =   "%PLAYERNAME% = The players name.                           %PLAYERCLASS% = The player's class."
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmQuestEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdOk_Click()
    QuestEditorOk
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub scrlGiveItem_Change()
    lblGiveITem.Caption = scrlGiveItem.Value
    If scrlGiveItem.Value > 0 Then
        lblGIveItemName.Caption = Trim$(Item(scrlGiveItem.Value).Name)
    Else
        lblGIveItemName.Caption = "None"
    End If
End Sub

Private Sub scrlStartNpc_Change()
    lblStartNpc.Caption = scrlStartNpc.Value
    If scrlStartNpc.Value > 0 Then
        lblStartNpcName.Caption = Trim$(Npc(scrlStartNpc.Value).Name)
    Else
        lblStartNpcName.Caption = "None"
    End If
End Sub

Private Sub scrlEndNpc_Change()
    lblEndNpc.Caption = scrlEndNpc.Value
    If scrlEndNpc.Value > 0 Then
        lblEndNpcName.Caption = Trim$(Npc(scrlEndNpc.Value).Name)
    Else
        lblEndNpcName.Caption = "None"
    End If
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = scrlLevel.Value
End Sub

Private Sub cmbQuestType_Click(Index As Integer)
Dim i As Long

    txtRequired(Index).Visible = True
    
    cmbNeedsIndex(Index).Clear
    cmbNeedsIndex(Index).AddItem "None"
    Select Case cmbQuestType(Index).ListIndex
        Case QuestTypes.KillNpc
            For i = 1 To MAX_NPCS
                cmbNeedsIndex(Index).AddItem i & ": " & Trim$(Npc(i).Name)
            Next
        Case QuestTypes.ItemCollection
            For i = 1 To MAX_ITEMS
                cmbNeedsIndex(Index).AddItem i & ": " & Trim$(Item(i).Name)
            Next
        Case QuestTypes.ExploreMap
            For i = 1 To MAX_MAPS
                cmbNeedsIndex(Index).AddItem i & ": " ' TODO: Add mapname & Trim$(Map(i).Name)
            Next
            txtRequired(Index).Text = 1
            txtRequired(Index).Visible = False
    End Select
    cmbNeedsIndex(Index).ListIndex = 0
End Sub

Private Sub scrlPart_Change()
Dim i As Long

    lblPart.Caption = scrlPart.Value
    For i = 1 To MAX_QUEST_NEEDS
        cmbQuestType(i).Visible = False
        cmbNeedsIndex(i).Visible = False
        txtRequired(i).Visible = False
    Next
    ' Now load up the data for this need
    cmbQuestType(scrlPart.Value).Visible = True
    cmbNeedsIndex(scrlPart.Value).Visible = True
    txtRequired(scrlPart.Value).Visible = True
    
    Select Case cmbQuestType(scrlPart.Value).ListIndex
        Case QuestTypes.KillNpc
        Case QuestTypes.ItemCollection
        Case QuestTypes.ExploreMap
            txtRequired(scrlPart.Value).Text = 1
            txtRequired(scrlPart.Value).Visible = False
    End Select
End Sub

Private Sub scrlPrvQuest_Change()
    lblPrevQuest.Caption = scrlPrvQuest.Value
    If scrlPrvQuest.Value > 0 Then
        lblPrevQuestName.Caption = Quest(scrlPrvQuest.Value).Name
    Else
        lblPrevQuestName.Caption = "None"
    End If
End Sub

Private Sub scrlRewardIndex_Change()
Dim i As Long

    lblRewardIndex.Caption = scrlRewardIndex.Value
    For i = 1 To MAX_QUEST_REWARDS
        cmbItemIndex(i).Visible = False
        txtItemValue(i).Visible = False
        chkSelectionOnly(i).Visible = False
    Next
    ' Now load up the data for this reward
    cmbItemIndex(scrlRewardIndex.Value).Visible = True
    txtItemValue(scrlRewardIndex.Value).Visible = True
    chkSelectionOnly(scrlRewardIndex.Value).Visible = True
End Sub


