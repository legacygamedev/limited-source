VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alert"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblOkay 
      BackColor       =   &H00000000&
      Caption         =   "Okay"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblYes 
      BackColor       =   &H00000000&
      Caption         =   "Yes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblNo 
      BackColor       =   &H00000000&
      Caption         =   "No"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AlertCallBack Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public OkayOnly As Boolean
Public sMessage As String
Public YesNo As Byte
Public callBack As Long

Private Sub Form_Load()
    If OkayOnly Then
        lblOkay.Visible = True
    Else
        lblYes.Visible = True
        lblNo.Visible = True
    End If
    lblMessage.Caption = sMessage
End Sub

Private Sub lblNo_Click()
    YesNo = NO
    ProcessClick
End Sub

Private Sub lblOkay_Click()
    YesNo = OKAY
    ProcessClick
End Sub

Private Sub lblYes_Click()
    YesNo = YES
    ProcessClick
End Sub

Private Sub ProcessClick()
    ' Make sure there is something to call back to
    If callBack <> 0 Then AlertCallBack callBack, YesNo, 0, 0, 0
    Unload Me
End Sub
