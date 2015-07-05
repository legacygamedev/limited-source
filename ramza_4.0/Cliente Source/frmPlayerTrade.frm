VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BorderStyle     =   0  'None
   Caption         =   "Trading"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerTrade.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Items2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ListBox Items1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox PlayerInv1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tus items para Comcerciar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tus Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   825
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remover"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar Comercio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5280
      TabIndex        =   5
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comerciar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comerciar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Packet As String
Dim i As Long

    Packet = "swapitems" & SEP_CHAR
    For i = 1 To MAX_PLAYER_TRADES
        Packet = Packet & Trading(i).InvNum & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendData(Packet)
    
    If Command1.ForeColor = &HFF& Then
        Command1.ForeColor = &HC000&
    Else
        Command1.ForeColor = &HFF&
    End If
End Sub

Private Sub Command3_Click()
Dim i As Long, n As Long
i = PlayerInv1.ListIndex + 1

If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
    For n = 1 To MAX_PLAYER_TRADES
        If Trading(n).InvNum = i Then
            MsgBox "Solo puedes comerciar ese item una vez!"
            Exit Sub
        End If
        If Trading(n).InvNum <= 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                MsgBox "No puedes comerciar dinero!"
                Exit Sub
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerBeltSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                    MsgBox "No puedes comerciar items equipados!"
                    Exit Sub
                Else
                    PlayerInv1.List(i - 1) = PlayerInv1.Text & " **"
                    Items1.List(n - 1) = n & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    Trading(n).InvNum = i
                    Trading(n).InvName = Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        End If
    Next n
End If
End Sub

Private Sub Command4_Click()
Dim i As Long, n As Long
i = Items1.ListIndex + 1

    If Trading(i).InvNum <= 0 Then
        MsgBox "No hay item para remover!"
        Exit Sub
    End If

    PlayerInv1.List(Trading(i).InvNum - 1) = Mid(Trim(PlayerInv1.List(Trading(i).InvNum - 1)), 1, Len(PlayerInv1.List(Trading(i).InvNum - 1)) - 3)
    Items1.List(i - 1) = n & ": <Nada>"
    Trading(i).InvNum = 0
    Trading(i).InvName = ""
    Call SendData("updatetradeinv" & SEP_CHAR & i & SEP_CHAR & 0 & SEP_CHAR & "" & SEP_CHAR & END_CHAR)
    Command1.ForeColor = &H80000012
End Sub

Private Sub Command5_Click()
    Call SendData("qtrade" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\Trade" & Ending) Then frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\Trade" & Ending)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command1.ForeColor = &H0&
    Command2.ForeColor = &H0&
End Sub

