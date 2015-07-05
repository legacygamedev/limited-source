VERSION 5.00
Begin VB.Form frmRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Editor V1 - Made by GIAKEN"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save changes"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Frame fraChar3 
      Caption         =   "Character 3"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   9135
      Begin VB.TextBox lblCSpeed 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   5460
         TabIndex        =   64
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox lblCDef 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3170
         TabIndex        =   58
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox lblCStr 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3170
         TabIndex        =   55
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox lblCHP 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   52
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCSprite 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3260
         TabIndex        =   40
         Top             =   340
         Width           =   1335
      End
      Begin VB.TextBox lblCClass 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   740
         TabIndex        =   34
         Top             =   1420
         Width           =   1335
      End
      Begin VB.CheckBox chkUsed 
         Caption         =   "Character Used?"
         Height          =   255
         Index           =   3
         Left            =   7440
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox lblCAccess 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   860
         TabIndex        =   24
         Top             =   1060
         Width           =   1335
      End
      Begin VB.TextBox lblCLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   750
         TabIndex        =   19
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   750
         TabIndex        =   13
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   4920
         TabIndex        =   61
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "DEF:"
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "STR:"
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "HP:"
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Class:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Access:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraChar2 
      Caption         =   "Character 2"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   9135
      Begin VB.TextBox lblCSpeed 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5460
         TabIndex        =   63
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox lblCDef 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3170
         TabIndex        =   57
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox lblCStr 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3170
         TabIndex        =   54
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox lblCHP 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   51
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCSprite 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3260
         TabIndex        =   39
         Top             =   340
         Width           =   1335
      End
      Begin VB.TextBox lblCClass 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   740
         TabIndex        =   33
         Top             =   1420
         Width           =   1335
      End
      Begin VB.CheckBox chkUsed 
         Caption         =   "Character Used?"
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblCAccess 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   860
         TabIndex        =   23
         Top             =   1060
         Width           =   1335
      End
      Begin VB.TextBox lblCLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   750
         TabIndex        =   18
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   750
         TabIndex        =   12
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   4920
         TabIndex        =   60
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "DEF:"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "STR:"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "HP:"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Class:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Access:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraChar1 
      Caption         =   "Character 1"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9135
      Begin VB.TextBox lblCSpeed 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5460
         TabIndex        =   62
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox lblCDef 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3170
         TabIndex        =   56
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox lblCStr 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3170
         TabIndex        =   53
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox lblCHP 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   50
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCSprite 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3260
         TabIndex        =   38
         Top             =   340
         Width           =   1335
      End
      Begin VB.TextBox lblCClass 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   740
         TabIndex        =   32
         Top             =   1420
         Width           =   1335
      End
      Begin VB.CheckBox chkUsed 
         Caption         =   "Character Used?"
         Height          =   255
         Index           =   1
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblCAccess 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   860
         TabIndex        =   25
         Top             =   1060
         Width           =   1335
      End
      Begin VB.TextBox lblCLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   750
         TabIndex        =   17
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox lblCName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   750
         TabIndex        =   9
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "This will NOT work if the character is logged in"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   66
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   $"frmRun.frx":058A
         Height          =   855
         Left            =   5280
         TabIndex        =   65
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label25 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   4920
         TabIndex        =   59
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "DEF:"
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "STR:"
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "HP:"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Class:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Access:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lblPlayerOnOff 
      Alignment       =   2  'Center
      Caption         =   "Note: Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label lblAccountPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Password: Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label lblAccountName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name: Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    'Me.Hide
    Call ClearMe
    'frmRun.Show

End Sub

Private Sub cmdDelete_Click()
Dim vbYesNo
    
    If Player(1).InGame = YES Then
        Call MsgBox("The player is online, cannot delete account.", vbOKOnly, "Ooops")
        Exit Sub
    End If

    If MsgBox("Are you 100% sure you want to delete this WHOLE account?", vbYesNo, "ARE YOU SURE?") = vbNo Then Exit Sub
    
    Call ClearPlayer(1)
    
    Kill (App.Path & "\accounts\" & Trim$(Player(1).Login) & ".ini")
    
    Call ClearMe

End Sub

Private Sub cmdSave_Click()

    Call DoSave

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

