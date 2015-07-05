VERSION 5.00
Begin VB.Form frmDrop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drop Amount"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtammount 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      TabIndex        =   3
      Text            =   "1"
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label cmdOk 
      BackStyle       =   0  'Transparent
      Caption         =   "Drop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1688
      TabIndex        =   4
      Top             =   945
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ammount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   465
      TabIndex        =   2
      Top             =   510
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   810
      TabIndex        =   1
      Top             =   165
      Width           =   540
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1395
      TabIndex        =   0
      Top             =   180
      Width           =   1890
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Ammount As Long

Private Sub Form_Load()
Dim InvNum As Long
Dim i As Long
Dim Ending As String

    InvNum = frmInventory.lstInv.ListIndex + 1
    
    frmDrop.lblName = Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
    Call ProcessAmmount
    
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\drop" & Ending) Then frmDrop.Picture = LoadPicture(App.Path & "\core files\interface\drop" & Ending)
    Next i
End Sub

Private Sub cmdOk_Click()
Dim InvNum As Long

    Ammount = txtammount.Text
    InvNum = frmInventory.lstInv.ListIndex + 1
    
    Call SendDropItem(InvNum, Ammount)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ProcessAmmount()
Dim InvNum As Long

    InvNum = frmInventory.lstInv.ListIndex + 1
        
    ' Check if more then max and set back to max if so
    If Ammount > GetPlayerInvItemValue(MyIndex, InvNum) Then
        Ammount = GetPlayerInvItemValue(MyIndex, InvNum)
    End If
    
    ' Make sure its not 0
    If Ammount <= 0 Then
        Ammount = 1
    End If
    

End Sub

Private Sub lblAmmount_Click()

End Sub

