VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   0  'None
   Caption         =   "Alert"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblYes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblOkay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OkayOnly As Boolean
Public sMessage As String
Public YesNo As Byte

Private Sub Form_Load()
    lblMessage.Caption = sMessage
    
    If OkayOnly Then
        frmAlert.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\MsgBox_Okay.bmp")
        frmAlert.lblOkay.Visible = True
        frmAlert.lblYes.Visible = False
        frmAlert.lblNo.Visible = False
    Else
        frmAlert.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\MsgBox_YesNo.bmp")
        frmAlert.lblOkay.Visible = False
        frmAlert.lblYes.Visible = True
        frmAlert.lblNo.Visible = True
    End If
End Sub

Private Sub lblNo_Click()
    YesNo = NO
    Audio.PlaySound ButtonBuzzer
    Unload Me
End Sub

Private Sub lblOkay_Click()
    Audio.PlaySound ButtonClick
    Unload Me
End Sub

Private Sub lblYes_Click()
    YesNo = YES
    Audio.PlaySound ButtonClick
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lblYes.Visible Then
            lblYes_Click
        Else
            lblOkay_Click
        End If
        KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyEscape Then
        If lblNo.Visible Then
            lblNo_Click
        End If
        KeyAscii = 0
    End If
End Sub
