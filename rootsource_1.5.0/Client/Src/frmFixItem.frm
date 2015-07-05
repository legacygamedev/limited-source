VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFixItem.frx":0000
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblFix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   4680
      UseMnemonic     =   0   'False
      Width           =   555
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   795
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Private Sub Form_Load()
    Me.Caption = GAME_NAME
    Me.Picture = LoadPicture(App.Path & "/gfx/interface/Menu.bmp")
End Sub

Private Sub lblFix_Click()
Dim Buffer As clsBuffer
    If (cmbItem.ListIndex + 1) > 0 Then
        If (cmbItem.ListIndex + 1) <= MAX_ITEMS Then
            Set Buffer = New clsBuffer
            Buffer.PreAllocate 6
            Buffer.WriteInteger CFixItem
            Buffer.WriteLong cmbItem.ListIndex + 1
            Call SendData(Buffer.ToArray())
        End If
    End If
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

