VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picCreditsMenu 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3000
      ScaleHeight     =   1815
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblCancelCredits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   42
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   $"frmMainMenu.frx":AFCC2
         Height          =   1095
         Left            =   240
         TabIndex        =   43
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox picNewAccountMenu 
      Height          =   1095
      Left            =   3360
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtPasswordNew 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "Password"
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtNameNew 
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Text            =   "Name"
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblCancelNew 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblConnectNew 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.PictureBox picDeleteAccountMenu 
      Height          =   1095
      Left            =   3360
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtPasswordDelete 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "Password"
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtNameDelete 
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Text            =   "Name"
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblCancelDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblConnectDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.PictureBox picLoginMenu 
      Height          =   1095
      Left            =   3360
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtPasswordLogin 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "Password"
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtNameLogin 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Text            =   "Name"
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblCancelLogin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblConnectLogin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connect"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.PictureBox picChars 
      Height          =   2175
      Left            =   3000
      ScaleHeight     =   2115
      ScaleWidth      =   5715
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox lstChars 
         Height          =   1230
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblCharsCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblDelChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete Char"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblUseChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Use Char"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblNewChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New Char"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCreateChar 
      Height          =   2175
      Left            =   3600
      ScaleHeight     =   2115
      ScaleWidth      =   4635
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtCharName 
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Text            =   "Name"
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   1080
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblDEX 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1920
         TabIndex        =   51
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label27 
         Caption         =   "Dexterity"
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Magic"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Speed"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Defence"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Strength"
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "SP"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "MP"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "HP"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblCreateCharCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblAddChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Char"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblMAGI 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblSPEED 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblDEF 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblSTR 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   90
      End
      Begin VB.Label lblSP 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblMP 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1830
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblDeleteAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

' *******************
' New Account Section
' *******************

Private Sub lblNewAccount_Click()
    picNewAccountMenu.Visible = True
End Sub

Private Sub lblCancelNew_Click()
    picNewAccountMenu.Visible = False
End Sub

Private Sub lblConnectNew_Click()
Dim Msg As String
Dim i As Long

    If Trim(txtNameNew.Text) <> "" And Trim(txtPasswordNew.Text) <> "" Then
        Msg = Trim(txtNameNew.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtNameNew.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

' **********************
' Delete Account Section
' **********************

Private Sub lblDeleteAccount_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        picDeleteAccountMenu.Visible = True
    End If
End Sub

Private Sub lblCancelDelete_Click()
    picDeleteAccountMenu.Visible = False
End Sub

Private Sub lblConnectDelete_Click()
    If Trim(txtNameDelete.Text) <> "" And Trim(txtPasswordDelete.Text) <> "" Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

' *************
' Login Section
' *************

Private Sub lblLogin_Click()
    picLoginMenu.Visible = True
End Sub

Private Sub lblCancelLogin_Click()
    picLoginMenu.Visible = False
End Sub

Private Sub lblConnectLogin_Click()
    If Trim(txtNameLogin.Text) <> "" And Trim(txtPasswordLogin.Text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

' *******
' Credits
' *******

Private Sub lblCredits_Click()
    picCreditsMenu.Visible = True
End Sub

Private Sub lblCancelCredits_Click()
    picCreditsMenu.Visible = False
End Sub

' ****
' Quit
' ****

Private Sub lblQuit_Click()
    Call GameDestroy
End Sub

' ***************************
' Character Selection Section
' ***************************

Private Sub lblCharsCancel_Click()
    Call TcpDestroy
    frmMainMenu.picChars.Visible = False
End Sub

Private Sub lblNewChar_Click()
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub lblUseChar_Click()
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub lblDelChar_Click()
Dim Value As Long

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

' **************************
' Character Creation Section
' **************************

Private Sub cmbClass_Click()
    frmMainMenu.lblHP.Caption = STR(Class(cmbClass.ListIndex).HP)
    frmMainMenu.lblMP.Caption = STR(Class(cmbClass.ListIndex).MP)
    frmMainMenu.lblSP.Caption = STR(Class(cmbClass.ListIndex).SP)
    
    frmMainMenu.lblSTR.Caption = STR(Class(cmbClass.ListIndex).STR)
    frmMainMenu.lblDEF.Caption = STR(Class(cmbClass.ListIndex).DEF)
    frmMainMenu.lblSPEED.Caption = STR(Class(cmbClass.ListIndex).SPEED)
    frmMainMenu.lblMAGI.Caption = STR(Class(cmbClass.ListIndex).MAGI)
End Sub

Private Sub lblAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim(frmMainMenu.txtCharName.Text) <> "" Then
        Msg = Trim(frmMainMenu.txtCharName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                frmMainMenu.txtCharName.Text = ""
                Exit Sub
            End If
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub lblCreateCharCancel_Click()
    frmMainMenu.picCreateChar.Visible = False
    frmMainMenu.picChars.Visible = True
End Sub

