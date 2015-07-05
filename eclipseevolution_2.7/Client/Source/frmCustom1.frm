VERSION 5.00
Begin VB.Form frmCustom1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   105
      ScaleHeight     =   553
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   788
      TabIndex        =   0
      Top             =   120
      Width           =   11820
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1140
         Top             =   6270
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   0
         Left            =   2790
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   0
         Left            =   750
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   150
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   2820
      End
   End
End
Attribute VB_Name = "frmCustom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If CUSTOM_IS_CLOSABLE = 1 Then
    Else
        Cancel = 1
        frmCustom1.Visible = True
    End If
    Timer1.Enabled = False
End Sub

Private Sub Form_LostFocus()
    On Error Resume Next
    If Me.Visible = True Then
        Me.SetFocus
    End If
End Sub

Private Sub picCustom_Click(Index As Integer)
    Dim packet As String
    Dim Custom_Type As Long
    Dim custom_string As String

    Custom_Type = 1
    custom_string = " "

    packet = "custommenuclick" & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & Custom_Type & SEP_CHAR & custom_string & END_CHAR
    Call SendData(packet)

End Sub

Private Sub Timer1_Timer()
    'On Error Resume Next
    If Me.Visible = True Then
        ' Me.SetFocus
        Call AlwaysOnTop(Me, True)
    End If
End Sub

Private Sub txtcustomOK_Click(Index As Integer)
    Dim packet As String
    Dim Custom_Type As Long
    Dim custom_string As String

    Custom_Type = 2
    custom_string = frmCustom1.txtCustom(Index).Text

    packet = "custommenuclick" & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & 2 & SEP_CHAR & custom_string & END_CHAR
    Call SendData(packet)

End Sub

Private Sub BtnCustom_Click(Index As Integer)

    Dim packet As String
    Dim Custom_Type As Long
    Dim custom_string As String

    Custom_Type = 3
    custom_string = " "

    packet = "custommenuclick" & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & Custom_Type & SEP_CHAR & custom_string & END_CHAR
    Call SendData(packet)

End Sub
