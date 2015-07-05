VERSION 5.00
Begin VB.Form frmBClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Block Class"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmBClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBlockedClasses 
      Caption         =   "Blocked Classes"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.HScrollBar scrlNum1 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   5
         Top             =   600
         Width           =   4335
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
         TabIndex        =   4
         Top             =   2280
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
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum2 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
      End
      Begin VB.HScrollBar scrlNum3 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label lblNum1 
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
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   75
      End
      Begin VB.Label lblAllowClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Class 1:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblAllowClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Class 2:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblAllowClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Class 3:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lblNum2 
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
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   75
      End
      Begin VB.Label lblNum3 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlNum1.Value = 0
    scrlNum2.Value = 0
    scrlNum3.Value = 0

    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblNum1.Caption = scrlNum1.Value & " - " & Trim$(Class(scrlNum1.Value).name)
    lblNum2.Caption = scrlNum2.Value & " - " & Trim$(Class(scrlNum2.Value).name)
    lblNum3.Caption = scrlNum3.Value & " - " & Trim$(Class(scrlNum3.Value).name)

    If EditorItemNum1 < scrlNum1.min Then
        EditorItemNum1 = scrlNum1.min
    End If

    scrlNum1.Value = EditorItemNum1

    If EditorItemNum2 < scrlNum2.min Then
        EditorItemNum2 = scrlNum2.min
    End If

    scrlNum2.Value = EditorItemNum2

    If EditorItemNum3 < scrlNum3.min Then
        EditorItemNum3 = scrlNum3.min
    End If

    scrlNum3.Value = EditorItemNum3
End Sub

Private Sub scrlNum1_Change()
    lblNum1.Caption = scrlNum1.Value & " - " & Trim$(Class(scrlNum1.Value).name)
    EditorItemNum1 = scrlNum1.Value
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = scrlNum2.Value & " - " & Trim$(Class(scrlNum2.Value).name)
    EditorItemNum2 = scrlNum2.Value
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = scrlNum3.Value & " - " & Trim$(Class(scrlNum3.Value).name)
    EditorItemNum3 = scrlNum3.Value
End Sub
