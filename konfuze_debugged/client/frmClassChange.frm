VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmClassChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class Change Attribute"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmClassChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
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
      TabCaption(0)   =   "Set Class"
      TabPicture(0)   =   "frmClassChange.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblClass"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblReqClass"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlClass"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOk"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlReqClass"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.HScrollBar scrlReqClass 
         Height          =   255
         Left            =   360
         Max             =   30
         Min             =   -1
         TabIndex        =   6
         Top             =   600
         Value           =   -1
         Width           =   4095
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
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
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar scrlClass 
         Height          =   255
         Left            =   360
         Max             =   30
         TabIndex        =   1
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblReqClass 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Req Class:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblClass 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Class:"
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
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmClassChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ClassChange = scrlClass.Value
    ClassChangeReq = scrlReqClass.Value
    Unload Me
End Sub

Private Sub Form_Load()
    If scrlReqClass.Value = -1 Then
        lblReqClass.Caption = scrlReqClass.Value & " - None"
    Else
        lblReqClass.Caption = scrlReqClass.Value & " - " & Trim(Class(scrlReqClass.Value).Name)
    End If
    lblClass.Caption = scrlClass.Value & " - " & Trim(Class(scrlClass.Value).Name)
    
    If ClassChange < scrlClass.Min Then ClassChange = scrlClass.Min
    scrlClass.Value = ClassChange
    If ClassChangeReq < scrlReqClass.Min Then ClassChangeReq = scrlReqClass.Min
    scrlReqClass.Value = ClassChangeReq
End Sub

Private Sub scrlClass_Change()
    lblClass.Caption = scrlClass.Value & " - " & Trim(Class(scrlClass.Value).Name)
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.Value = -1 Then
        lblReqClass.Caption = scrlReqClass.Value & " - None"
    Else
        lblReqClass.Caption = scrlReqClass.Value & " - " & Trim(Class(scrlReqClass.Value).Name)
    End If
End Sub
