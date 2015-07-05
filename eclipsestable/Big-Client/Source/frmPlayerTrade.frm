VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmPlayerTrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
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
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   2655
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
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
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3120
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   2445
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3120
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim Packet As String
    Dim i As Long

    Packet = "swapitems" & SEP_CHAR
    For i = 1 To MAX_PLAYER_TRADES
        Packet = Packet & Trading(i).InvNum & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendData(Packet)

    If Command1.ForeColor = &HFF00& Then
        Command1.ForeColor = &H0&
    Else
        Command1.ForeColor = &HFF00&
    End If
End Sub

Private Sub Command3_Click()
    Dim i As Long
    Dim n As Long
    Dim GoldAmount As Long
    
    i = PlayerInv1.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        For n = 1 To MAX_PLAYER_TRADES
            If Trading(n).InvNum = i Then
                MsgBox "You can only trade that item once!"
                Exit Sub
            End If
            If Trading(n).InvNum <= 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Bound = 1 Then
            Call AddText("This item is not able to be traded, it is probably a quest item or an event item.", 12)
            Exit Sub
            End If
                If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            GoldAmount = Val(InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & "(" & GetPlayerInvItemValue(MyIndex, i) & ") would you like to trade?", "Trade " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name), "", frmPlayerTrade.Left, frmPlayerTrade.top))
            If IsNumeric(GoldAmount) Then
            If Int(GoldAmount) > Int(GetPlayerInvItemValue(MyIndex, i)) Then
            Call AddText("You don't have that amount to trade!", BRIGHTRED)
            Exit Sub
            End If
                        PlayerInv1.List(i - 1) = PlayerInv1.Text & " **"
                        Items1.List(n - 1) = n & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " / " & GoldAmount & ""
                        Trading(n).InvNum = i
                        Trading(n).InvName = "" & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & ""
                        Trading(n).InvVal = "" & Int(GoldAmount) & ""
                Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & SEP_CHAR & GoldAmount & END_CHAR)
            Exit Sub
            Else
            Exit Sub
            End If
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                        MsgBox "Cant trade worn items!"
                        Exit Sub
                    Else
                        PlayerInv1.List(i - 1) = PlayerInv1.Text & " **"
                        Items1.List(n - 1) = n & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                        Trading(n).InvNum = i
                        Trading(n).InvName = Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                        Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & SEP_CHAR & 0 & END_CHAR)
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
        MsgBox "No item to remove!"
        Exit Sub
    End If

    PlayerInv1.List(Trading(i).InvNum - 1) = Mid$(Trim$(PlayerInv1.List(Trading(i).InvNum - 1)), 1, Len(PlayerInv1.List(Trading(i).InvNum - 1)) - 3)
    Items1.List(i - 1) = n & ": <Nothing>"
    Trading(i).InvNum = 0
    Trading(i).InvName = vbNullString
    Trading(i).InvVal = 0
    Call SendData("updatetradeinv" & SEP_CHAR & i & SEP_CHAR & 0 & SEP_CHAR & vbNullString & SEP_CHAR & 0 & END_CHAR)
    Command1.ForeColor = &H80000012
End Sub


Private Sub Command5_Click()
    Call SendData("qtrade" & END_CHAR)
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim Ending As String
    For i = 1 To 3
        If i = 1 Then
            Ending = ".gif"
        End If
        If i = 2 Then
            Ending = ".jpg"
        End If
        If i = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Trade" & Ending) Then
            frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\Trade" & Ending)
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command1.ForeColor = &H0&
    Command2.ForeColor = &H0&
End Sub

