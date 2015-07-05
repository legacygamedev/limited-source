VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7515
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
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblFix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Item"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the item you wish to fix."
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
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
            
            Buffer.WriteLong CFixItem
            Buffer.WriteLong cmbItem.ListIndex + 1
            
            SendData Buffer.ToArray()
            
            Set Buffer = Nothing
        
        End If
    End If
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

