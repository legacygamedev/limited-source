VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Item"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
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
   Icon            =   "frmMapItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info."
      TabPicture(0)   =   "frmMapItem.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5175
         Begin VB.TextBox txtValue 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   9
            Text            =   "1"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox cmbItem 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMapItem.frx":0E5E
            Left            =   120
            List            =   "frmMapItem.frx":0E71
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   4935
         End
         Begin VB.CheckBox chkItemBlocked 
            Caption         =   "Blocked?"
            Height          =   270
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   840
            Max             =   100
            Min             =   1
            TabIndex        =   4
            Top             =   960
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMapItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
    
    cmbItem.Clear
    For i = 1 To MAX_ITEMS
        cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next
    cmbItem.ListIndex = 0
End Sub

Private Sub cmdOk_Click()
    ItemEditorNum = cmbItem.ListIndex + 1
    ItemEditorValue = scrlValue.Value
    ItemEditorBlocked = chkItemBlocked.Value
    Unload Me
End Sub

Private Sub scrlValue_Change()
    txtValue.Text = scrlValue.Value
End Sub

Private Sub txtValue_Change()
    SetTextBox txtValue, scrlValue
End Sub
