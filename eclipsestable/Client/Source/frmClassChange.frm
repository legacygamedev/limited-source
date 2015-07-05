VERSION 5.00
Begin VB.Form frmClassChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class Change Attribute"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmClassChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraSetClass 
      Caption         =   "Set Class"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.HScrollBar scrlClass 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   4
         Top             =   1080
         Width           =   4095
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
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
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
         Left            =   2520
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
      Begin VB.HScrollBar scrlReqClass 
         Height          =   255
         Left            =   240
         Max             =   30
         Min             =   -1
         TabIndex        =   1
         Top             =   480
         Value           =   -1
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   375
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
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
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
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmClassChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        lblReqClass.Caption = scrlReqClass.Value & " - " & Trim$(Class(scrlReqClass.Value).name)
    End If
    lblClass.Caption = scrlClass.Value & " - " & Trim$(Class(scrlClass.Value).name)

    If ClassChange < scrlClass.min Then
        ClassChange = scrlClass.min
    End If
    scrlClass.Value = ClassChange
    If ClassChangeReq < scrlReqClass.min Then
        ClassChangeReq = scrlReqClass.min
    End If
    scrlReqClass.Value = ClassChangeReq
End Sub


Private Sub scrlClass_Change()
    lblClass.Caption = scrlClass.Value & " - " & Trim$(Class(scrlClass.Value).name)
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.Value = -1 Then
        lblReqClass.Caption = scrlReqClass.Value & " - None"
    Else
        lblReqClass.Caption = scrlReqClass.Value & " - " & Trim$(Class(scrlReqClass.Value).name)
    End If
End Sub
