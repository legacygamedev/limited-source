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
      Height          =   390
      Left            =   0
      Picture         =   "frmTrade.frx":586A4
      ScaleHeight     =   390
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   3990
      Width           =   1065
   End
   Begin VB.PictureBox picFixItems 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   30
      Picture         =   "frmTrade.frx":59CD6
      ScaleHeight     =   390
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   3105
      Width           =   1065
   End
   Begin VB.PictureBox picDeal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "frmTrade.frx":5B308
      ScaleHeight     =   390
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   3540
      Width           =   1065
   End
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2970
      ItemData        =   "frmTrade.frx":5C93A
      Left            =   75
      List            =   "frmTrade.frx":5C93C
      TabIndex        =   0
      Top             =   75
      Width           =   3495
   End
   Begin VB.Label lblRestock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESTOCK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1395
      TabIndex        =   12
      Top             =   3540
      Width           =   1860
   End
   Begin VB.Label lblHSPD 
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   11
      Top             =   3960
      Width           =   1950
   End
   Begin VB.Label lblHMAG 
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   10
      Top             =   3600
      Width           =   1950
   End
   Begin VB.Label lblHDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   9
      Top             =   3780
      Width           =   1950
   End
   Begin VB.Label lblHSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   8
      Top             =   3420
      Width           =   1950
   End
   Begin VB.Label lblStock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1395
      TabIndex        =   7
      Top             =   3870
      Width           =   1860
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CostName: Cost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3870
      TabIndex        =   6
      Top             =   2640
      Width           =   2070
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   3930
      TabIndex        =   5
      Top             =   1650
      Width           =   1950
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3735
      TabIndex        =   4
      Top             =   900
      Width           =   2325
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
