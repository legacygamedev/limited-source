VERSION 5.00
Begin VB.Form frmTrade 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6195
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTrade.frx":0000
   ScaleHeight     =   4380
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmTrade.frx":586A2
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   3885
      Width           =   1500
   End
   Begin VB.PictureBox picFixItems 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmTrade.frx":5A430
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   3495
      Width           =   1500
   End
   Begin VB.PictureBox picDeal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmTrade.frx":5C1BE
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   3105
      Width           =   1500
   End
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5763F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0023A0D6&
      Height          =   2970
      ItemData        =   "frmTrade.frx":5DF4C
      Left            =   120
      List            =   "frmTrade.frx":5DF4E
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblRestock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESTOCK"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblHSPD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblHMAG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblHDEF 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblHSTR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblStock 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblCost 
      BackStyle       =   0  'Transparent
      Caption         =   "CostName: Cost"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

Private Sub picDeal_Click()
    If lstTrade.ListCount > 0 Then
        Call SendData("traderequest" & SEP_CHAR & lstTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub picFixItems_Click()
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmFixItem.cmbItem.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmFixItem.cmbItem.AddItem "Unused Slot"
        End If
    Next i
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub

Private Sub lstTrade_Click()
    If lstTrade.ListCount > -1 Then
        Call SendData("tradegetitem" & SEP_CHAR & lstTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub
