VERSION 5.00
Begin VB.Form frmMinimap 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Minimap"
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLoad9 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   18
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   17
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   16
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   15
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   10
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   14
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   13
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   12
      Top             =   240
      Width           =   9585
   End
   Begin VB.PictureBox picLoad2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   5160
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   11
      Top             =   240
      Width           =   9585
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "User32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2
Public Function Transparent(Form As Form, Layout As Byte) As Boolean
    SetWindowLong Form.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Form.hWnd, 0, Layout, LWA_ALPHA
    Transparent = Err.LastDllError = 0
End Function

Private Sub Form_Load()
Dim sDc As Long

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
         Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmMinimap.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
      End Sub

Private Sub Label1_Click()
Unload Me
End Sub
