VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Selection"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
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
      Height          =   1830
      ItemData        =   "frmChars.frx":0FC2
      Left            =   240
      List            =   "frmChars.frx":0FC4
      TabIndex        =   0
      Top             =   1320
      Width           =   5565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2100
      TabIndex        =   1
      Top             =   3240
      Width           =   1800
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\CharacterSelect" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect" & Ending)
    Next i
End Sub

Private Sub Label1_Click()
    If AutoLogin = 1 Then
        Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))
        frmChars.Label1.Visible = False
        frmMainMenu.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
Dim BMU As BitmapUtils
Dim strfilename As String

    If lstChars.List(lstChars.ListIndex) <> "Free Character Slot" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If
    
            If ENCRYPT_TYPE = "BMP" Then
                frmNewChar.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "sprites" & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmNewChar.picSprites.Width = BMU.ImageWidth
                frmNewChar.picSprites.Height = BMU.ImageHeight
                Call BMU.Blt(frmNewChar.picSprites.hDC)
                End If
    
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
Dim BMU As BitmapUtils
Dim strfilename As String

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If
    
        If ENCRYPT_TYPE = "BMP" Then
            frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
            Else
            Set BMU = New BitmapUtils
            strfilename = App.Path & "/gfx/" & "items" & "." & Trim$(ENCRYPT_TYPE)
            BMU.LoadByteData (strfilename)
            BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
            BMU.DecompressByteData_ZLib
            frmMirage.picItems.Width = BMU.ImageWidth
            frmMirage.picItems.Height = BMU.ImageHeight
            Call BMU.Blt(frmMirage.picItems.hDC)
            End If
    
        If ENCRYPT_TYPE = "BMP" Then
                frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "sprites" & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmSpriteChange.picSprites.Width = BMU.ImageWidth
                frmSpriteChange.picSprites.Height = BMU.ImageHeight
                Call BMU.Blt(frmSpriteChange.picSprites.hDC)
                End If
            
    Call MenuState(MENU_STATE_USECHAR)
    
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

