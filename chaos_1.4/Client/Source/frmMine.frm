VERSION 5.00
Begin VB.Form frmMine 
   Caption         =   "Mining Skill"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.HScrollBar scrlToolNo 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.HScrollBar scrlOreNo 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtToolName 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtOreName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "The item number that the tool is stored under.)"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "(The name of the tool needed by the users to catch a ore from this location.)"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "(The number of the item that you want to be obtained as Ore.)"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "(The name of the Ore Obtained from this area.)"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Item Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Tool Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Item Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Ore Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tool Needed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ore:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OreName = frmMine.txtOreName.Text
ToolName = frmMine.txtToolName.Text
OreNumber = frmMine.scrlOreNo.Value
ToolNumber = frmMine.scrlToolNo.Value
Unload Me
End Sub

Private Sub scrlOreNo_Change()
Label9.Caption = scrlOreNo.Value
End Sub

Private Sub scrlToolNo_Change()
Label10.Caption = scrlToolNo.Value
End Sub
