VERSION 5.00
Begin VB.Form frmHouseEditor 
   Caption         =   "House Editor"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdED 
      Caption         =   "Eye Dropper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdtype 
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdtype 
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Map Grid "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   6465
      LargeChange     =   10
      Left            =   120
      Max             =   512
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6480
      Left            =   480
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   0
      Top             =   840
      Width           =   6720
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6480
         Left            =   0
         ScaleHeight     =   432
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   448
         TabIndex        =   1
         Top             =   0
         Width           =   6720
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4080
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   3105
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   2520
      Begin VB.OptionButton optF2Anim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   1080
      End
      Begin VB.OptionButton optFringe2 
         Caption         =   "Fringe 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.OptionButton optM2Anim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1245
      End
      Begin VB.OptionButton optMask2 
         Caption         =   "Mask 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame frmtile 
      Caption         =   "Tile Sheet"
      Height          =   735
      Left            =   2760
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmHouseEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KeyShift As Boolean

Private Sub cmdED_Click()
    If Me.MousePointer = 2 Or frmStable.MousePointer = 2 Then
        Me.MousePointer = 1
        frmStable.MousePointer = 1
    Else
        Me.MousePointer = 2
        frmStable.MousePointer = 2
    End If
End Sub

Private Sub cmdExit_Click()
    Dim X As Long

    X = MsgBox("Are you sure you want to discard your changes?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call HouseEditorCancel
End Sub

Private Sub cmdGrid_Click()
    If GridMode = 0 Then
        GridMode = 1
    Else
        GridMode = 0
    End If
End Sub

Private Sub cmdSave_Click()
    Dim X As Long

    X = MsgBox("Are you sure you want to make these changes?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call HouseEditorSend
End Sub

Private Sub cmdtype_Click(Index As Integer)
    If Index = 1 Then
        MapEditorSelectedType = 1

        Me.fraAttribs.Visible = False
        Me.fraLayers.Visible = True
        Me.frmtile.Visible = True
    ElseIf Index = 2 Then
        HouseEditorSelectedType = 2

        Me.shpSelected.Width = 32
        Me.shpSelected.Height = 32

        Me.fraLayers.Visible = False
        Me.frmtile.Visible = False
        Me.fraAttribs.Visible = True

    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True

    Me.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles9.bmp")

    EditorSet = 9

    scrlPicture.max = Int((picBackSelect.Height - picBack.Height) / PIC_Y)
End Sub

Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call HouseEditorChooseTile(Button, Shift, X, y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
            EditorTileY = Int(y / PIC_Y)

            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If

            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If

    If HouseEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int((shpSelected.Left + PIC_X) / PIC_X)
    EditorTileY = Int((shpSelected.Top + PIC_Y) / PIC_Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call HouseEditorChooseTile(Button, Shift, X, y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
            EditorTileY = Int(y / PIC_Y)

            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If

            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If

    If HouseEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.Top / PIC_Y)
End Sub

Private Sub scrlPicture_Change()
    Call HouseEditorTileScroll
End Sub
