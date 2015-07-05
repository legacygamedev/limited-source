VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEditSkills 
   Caption         =   "Edit skill"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   LinkTopic       =   "Form2"
   ScaleHeight     =   6600
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
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
      TabCaption(0)   =   "Edit skill"
      TabPicture(0)   =   "frmEditSkills.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOk"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAction"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFail"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSucces"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "HScroll1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "HScroll2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Picture1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   240
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   18
         Top             =   3840
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   19
            Top             =   15
            Width           =   480
            Begin VB.PictureBox iconn 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   20
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5400
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox txtSucces 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtFail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtAction 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
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
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1800
         TabIndex        =   2
         Top             =   6000
         Width           =   1215
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
         TabIndex        =   1
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   2760
         TabIndex        =   17
         Top             =   5400
         Width           =   75
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Left            =   2760
         TabIndex        =   16
         Top             =   4800
         Width           =   75
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
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
         TabIndex        =   13
         Top             =   5160
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Icon:"
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
         Top             =   3600
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Top:"
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
         TabIndex        =   11
         Top             =   4560
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "On Succes message:"
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
         TabIndex        =   9
         Top             =   2760
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "On fail message:"
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
         TabIndex        =   7
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "On action message:"
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
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
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
         TabIndex        =   4
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmEditSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim val1 As Long
Dim val2 As Long

Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 5
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExist("GFX\Icons" & Ending) Then iconn.Picture = LoadPicture(App.Path & "\GFX\Icons" & Ending)
    Next i

iconn.Left = Val(HScroll1.Value * PIC_X)
iconn.Top = Val(HScroll2.Value * PIC_Y)
val1 = 0
val2 = 0
End Sub

Private Sub HScroll1_Change()
iconn.Left = -Val(HScroll1.Value * PIC_X)
iconn.Top = -Val(HScroll2.Value * PIC_Y)
Label8.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
iconn.Left = -Val(HScroll1.Value * PIC_X)
iconn.Top = -Val(HScroll2.Value * PIC_Y)
Label9.Caption = HScroll2.Value
End Sub
