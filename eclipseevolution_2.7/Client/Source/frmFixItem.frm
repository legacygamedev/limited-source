VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fix Item"
   ClientHeight    =   5985
   ClientLeft      =   135
   ClientTop       =   315
   ClientWidth     =   5985
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
   Icon            =   "frmFixItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   2805
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   2
      Top             =   5520
      Width           =   5505
   End
   Begin VB.Label picFix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   1
      Top             =   5040
      Width           =   2805
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picFix_Click()
    Call SendData("fixitem" & SEP_CHAR & snumber & SEP_CHAR & cmbItem.ListIndex + 1 & END_CHAR)
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

        If FileExists("GUI\FixItems" & Ending) Then
            frmFixItem.Picture = LoadPicture(App.Path & "\GUI\FixItems" & Ending)
        End If
    Next i

    frmFixItem.cmbItem.Clear
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmFixItem.cmbItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmFixItem.cmbItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmFixItem.cmbItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmFixItem.cmbItem.addItem i & "> Empty"
        End If
    Next i

    frmFixItem.cmbItem.ListIndex = 0
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub
