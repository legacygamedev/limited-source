VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDrop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drop Item"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
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
   ScaleHeight     =   3360
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   5530
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
      TabCaption(0)   =   "Drop Panel"
      TabPicture(0)   =   "frmDrop.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAmmount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPlus1000"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPlus100"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPlus10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPlus1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOk"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCancel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdMinus1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdMinus10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdMinus100"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdMinus1000"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.CommandButton cmdMinus1000 
         Caption         =   "-1000"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdMinus100 
         Caption         =   "-100"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdMinus10 
         Caption         =   "-10"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdMinus1 
         Caption         =   "-1"
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
         TabIndex        =   1
         Top             =   1800
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
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   2640
         Width           =   2535
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
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdPlus1 
         Caption         =   "+1"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlus10 
         Caption         =   "+10"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlus100 
         Caption         =   "+100"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlus1000 
         Caption         =   "+1000"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ammount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblAmmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Ammount As Long

Private Sub Form_Load()
Dim InvNum As Long

    Ammount = 1
    InvNum = frmMirage.lstInv.ListIndex + 1
    
    frmDrop.lblName = Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
    Call ProcessAmmount
End Sub

Private Sub cmdOk_Click()
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1
    
    Call SendDropItem(InvNum, Ammount)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPlus1_Click()
    Ammount = Ammount + 1
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1_Click()
    Ammount = Ammount - 1
    Call ProcessAmmount
End Sub

Private Sub cmdPlus10_Click()
    Ammount = Ammount + 10
    Call ProcessAmmount
End Sub

Private Sub cmdMinus10_Click()
    Ammount = Ammount - 10
    Call ProcessAmmount
End Sub

Private Sub cmdPlus100_Click()
    Ammount = Ammount + 100
    Call ProcessAmmount
End Sub

Private Sub cmdMinus100_Click()
    Ammount = Ammount - 100
    Call ProcessAmmount
End Sub

Private Sub cmdPlus1000_Click()
    Ammount = Ammount + 1000
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1000_Click()
    Ammount = Ammount - 1000
    Call ProcessAmmount
End Sub

Private Sub ProcessAmmount()
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1
        
    ' Check if more then max and set back to max if so
    If Ammount > GetPlayerInvItemValue(MyIndex, InvNum) Then
        Ammount = GetPlayerInvItemValue(MyIndex, InvNum)
    End If
    
    ' Make sure its not 0
    If Ammount <= 0 Then
        Ammount = 1
    End If

    frmDrop.lblAmmount.Caption = Ammount & "/" & GetPlayerInvItemValue(MyIndex, InvNum)
End Sub

