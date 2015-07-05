VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (New Character)"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
   End
   Begin VB.PictureBox picinit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   7320
      Width           =   1005
   End
   Begin VB.Timer timCharPic 
      Interval        =   500
      Left            =   -15
      Top             =   0
   End
   Begin VB.PictureBox picChars 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1.21470e5
      Left            =   120
      ScaleHeight     =   1.21440e5
      ScaleWidth      =   5760
      TabIndex        =   14
      Top             =   8400
      Width           =   5790
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3600
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      ItemData        =   "frmNewChar.frx":0000
      Left            =   2160
      List            =   "frmNewChar.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.OptionButton optMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   3120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image picFemaleOn 
      Height          =   480
      Left            =   240
      Picture         =   "frmNewChar.frx":0004
      Top             =   7080
      Width           =   1320
   End
   Begin VB.Image picMaleOff 
      Height          =   480
      Left            =   1800
      Picture         =   "frmNewChar.frx":04EC
      Top             =   7800
      Width           =   1320
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   6000
      Picture         =   "frmNewChar.frx":099A
      Top             =   5640
      Width           =   1320
   End
   Begin VB.Image picAddChar 
      Height          =   480
      Left            =   360
      Picture         =   "frmNewChar.frx":0E7A
      Top             =   5760
      Width           =   1320
   End
   Begin VB.Image picFemale 
      Height          =   480
      Left            =   480
      Picture         =   "frmNewChar.frx":1347
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Image picMale 
      Height          =   480
      Left            =   480
      Picture         =   "frmNewChar.frx":182F
      Top             =   3480
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   1035
      Left            =   3555
      Top             =   5130
      Width           =   1065
   End
   Begin VB.Image picForward 
      Height          =   480
      Left            =   4680
      Picture         =   "frmNewChar.frx":1CDD
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image picBack 
      Height          =   480
      Left            =   3000
      Picture         =   "frmNewChar.frx":1FA5
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SP:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cha:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Wiz:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Con:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dex:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Int:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Str:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
   Begin VB.Image picMaleOn 
      Height          =   480
      Left            =   1800
      Picture         =   "frmNewChar.frx":226C
      Top             =   7080
      Width           =   1320
   End
   Begin VB.Image picFemaleOff 
      Height          =   480
      Left            =   240
      Picture         =   "frmNewChar.frx":271A
      Top             =   7800
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmNewChar.frx":2C02
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmNewChar.frx":4D75
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      Height          =   450
      Left            =   2130
      Top             =   3090
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblSpriteNo 
      Caption         =   "0"
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lblCHA 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblWIZ 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblINT 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblDEX 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCON 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmNewChar.frx":6F1B
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmNewChar.frx":9AEA
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngCounter As Long
Private lngNumber As Long


Private MoveArr() As String
Private moveStr As String

Dim spriteArr() As String
Dim spriteStrM As String
Dim spriteStrF As String

Private Sub cmbClass_Click()
    lblHP.Caption = str(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = str(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = str(Class(cmbClass.ListIndex).SP)
    
    lblSTR.Caption = str(Class(cmbClass.ListIndex).str)
    lblCha.Caption = str(Class(cmbClass.ListIndex).cha)
    lblINT.Caption = str(Class(cmbClass.ListIndex).intel)
    lblDex.Caption = str(Class(cmbClass.ListIndex).dex)
    lblCon.Caption = str(Class(cmbClass.ListIndex).con)
    lblWiz.Caption = str(Class(cmbClass.ListIndex).wiz)
End Sub






Private Sub Form_Load()
picChars.Picture = LoadPicture(App.Path & "\data\bmp\Sprites.bmp")
lngCounter = 12
moveStr = "0,1,0,1,0,1,0,1,0,1,0,1,3,4,3,4,3,4,3,4,3,4,3,4,6,7,6,7,6,7,6,7,6,7,6,7,9,10,9,10,9,10,9,10,9,10,9,10"
MoveArr = Split(moveStr, ",")
spriteStrM = "0,1,6,7,14,16,17,19,23,25,36,43,53,69,108,109,110,156,172,225"
spriteStrF = "32,33,34,35,42,44,54,66,81,82,83,151,164,168,190"
spriteArr = Split(spriteStrM, ",")
End Sub

Private Sub optFemale_Click()
spriteArr = Split(spriteStrF, ",")
lngNumber = 0
End Sub

Private Sub optMale_Click()
spriteArr = Split(spriteStrM, ",")
lngNumber = 0
End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long


        
    txtName.text = LCase(txtName.text)
    If Trim(txtName.text) <> "" Then
        Msg = Trim(txtName.text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) <> 95 Then
                If Asc(Mid(Msg, i, 1)) < 97 Or Asc(Mid(Msg, i, 1)) > 123 Then
                    Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                    txtName.text = ""
                    Exit Sub
                End If
            End If
        Next i
        Dim strName As String
        Dim underscoreCount As Long
        Dim blnNextUC As Boolean
        blnNextUC = True
        underscoreCount = 0
        strName = ""
        For i = 1 To Len(txtName.text)
            If blnNextUC Then
                strName = strName & UCase(Mid(txtName.text, i, 1))
                blnNextUC = False
            Else
                If Asc(Mid(Msg, i, 1)) = 95 Then
                    blnNextUC = True
                    underscoreCount = underscoreCount + 1
                End If
                If underscoreCount < 2 Then
                    strName = strName & Mid(txtName.text, i, 1)
                Else
                    Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                    txtName.text = ""
                End If
            End If
            
        Next i
        txtName.text = strName
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picBack_Click()
    lngNumber = lngNumber - 1
    If lngNumber < 0 Then lngNumber = 0
    lngCounter = 12
    Call CharGenBltSprite(spriteArr(lngNumber), Val(MoveArr(lngCounter)))
    lblSpriteNo.Caption = spriteArr(lngNumber)
    txtName.SetFocus
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub picFemale_Click()
    optFemale.value = True
    picFemale.Picture = picFemaleOn.Picture
    optMale.value = False
    picMale.Picture = picMaleOff.Picture
End Sub

Private Sub picForward_Click()
    lngNumber = lngNumber + 1
    If lngNumber > UBound(spriteArr) Then lngNumber = UBound(spriteArr)
    lngCounter = 12
    Call CharGenBltSprite(spriteArr(lngNumber), Val(MoveArr(lngCounter)))
    lblSpriteNo.Caption = spriteArr(lngNumber)
    txtName.SetFocus
End Sub

Private Sub picMale_Click()
    optFemale.value = False
    picFemale.Picture = picFemaleOff.Picture
    optMale.value = True
    picMale.Picture = picMaleOn.Picture
End Sub

Private Sub picPath_Click(Index As Integer)
If Index > 5 Then Index = 5
    lblHP.Caption = str(Class(Index).HP)
    lblMP.Caption = str(Class(Index).MP)
    lblSP.Caption = str(Class(Index).SP)
    
    lblSTR.Caption = str(Class(Index).str)
    lblCha.Caption = str(Class(Index).cha)
    lblINT.Caption = str(Class(Index).intel)
    lblDex.Caption = str(Class(Index).dex)
    lblCon.Caption = str(Class(Index).con)
    lblWiz.Caption = str(Class(Index).wiz)
    
    cmbClass.ListIndex = Index
        txtName.SetFocus
End Sub

Private Sub timCharPic_Timer()

    Call CharGenBltSprite(spriteArr(lngNumber), Val(MoveArr(lngCounter)))
    lblSpriteNo.Caption = spriteArr(lngNumber)
    lngCounter = lngCounter + 1
    
    If lngCounter >= UBound(MoveArr) Then
        lngCounter = 0
    End If
End Sub

