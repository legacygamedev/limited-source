VERSION 5.00
Begin VB.Form frmAttrib 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   103
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
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
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
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
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.HScrollBar scrlData 
      Height          =   255
      Index           =   3
      Left            =   1080
      Max             =   1
      Min             =   1
      TabIndex        =   5
      Top             =   800
      Value           =   1
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.HScrollBar scrlData 
      Height          =   255
      Index           =   2
      Left            =   1080
      Max             =   1
      Min             =   1
      TabIndex        =   4
      Top             =   450
      Value           =   1
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.HScrollBar scrlData 
      Height          =   255
      Index           =   1
      Left            =   1080
      Max             =   1
      Min             =   1
      TabIndex        =   3
      Top             =   95
      Value           =   1
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
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
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
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
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblAttribName 
      Caption         =   "AttribName3"
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
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAttribName 
      Caption         =   "AttribName2"
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
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAttribName 
      Caption         =   "AttribName1"
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
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    ClearMapAttribs
    Unload Me
    
End Sub

Private Sub cmdConfirm_Click()
Dim LoopI As Long

    For LoopI = 1 To 3
        If scrlData(LoopI).Visible Then MapAttribData(LoopI) = scrlData(LoopI).Value
    Next
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim LoopI As Long
Dim LastIndex As Long

    Me.Caption = MapAttribFormTitle
    
    For LoopI = 1 To 3
        If LenB(Trim$(MapAttribName(LoopI))) > 0 Then
            lblAttribName(LoopI) = MapAttribName(LoopI)
            
            scrlData(LoopI).Min = MapAttribMin(LoopI)
            scrlData(LoopI).Max = MapAttribMax(LoopI)
            If scrlData(LoopI).Min < 1 Then scrlData(LoopI).Value = 0 Else scrlData(LoopI).Value = 1
            
            lblAttribName(LoopI).Visible = True
            scrlData(LoopI).Visible = True
            lblResult(LoopI).Visible = True
            LastIndex = LoopI
        End If
    Next
    
    If LastIndex > 0 Then
        cmdConfirm.Top = lblResult(LastIndex).Top + 30
        cmdCancel.Top = cmdConfirm.Top
        Me.Height = (cmdConfirm.Top + 50) * Screen.TwipsPerPixelY
    End If
    
    Me.Show
    
End Sub

Private Sub scrlData_Change(Index As Integer)
    lblResult(Index) = scrlData(Index)
End Sub

Private Sub scrlData_Scroll(Index As Integer)
    scrlData_Change Index
End Sub
