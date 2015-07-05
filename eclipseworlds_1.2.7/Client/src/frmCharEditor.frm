VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCharEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Editor"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   652
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameSearch 
      Caption         =   "Players List"
      Height          =   5475
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   2685
      Begin MSComctlLib.ListView listCharacters 
         Height          =   4725
         Left            =   60
         TabIndex        =   5
         Top             =   660
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   8334
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   2
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Character Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   2565
      End
   End
   Begin VB.Frame frameCharPanel 
      Caption         =   "Editing - None"
      Height          =   5565
      Left            =   2760
      TabIndex        =   0
      Top             =   90
      Width           =   6945
      Begin VB.ComboBox cmbAccess 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCharEditor.frx":038A
         Left            =   3390
         List            =   "frmCharEditor.frx":03A0
         TabIndex        =   48
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton cmdFastUpdate 
         BackColor       =   &H0080FF80&
         Caption         =   "Fast Update"
         Height          =   315
         Left            =   120
         MaskColor       =   &H0080FF80&
         TabIndex        =   47
         Top             =   4140
         Width           =   1185
      End
      Begin VB.CommandButton cmdFastCancel 
         Caption         =   "Fast Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox txtSprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         TabIndex        =   44
         Text            =   "0"
         Top             =   2580
         Width           =   570
      End
      Begin MSComCtl2.UpDown upSprite 
         Height          =   555
         Left            =   4020
         TabIndex        =   43
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   979
         _Version        =   393216
         BuddyControl    =   "txtSprite"
         BuddyDispid     =   196615
         OrigLeft        =   3990
         OrigTop         =   1770
         OrigRight       =   4245
         OrigBottom      =   2265
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPoints 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   570
         TabIndex        =   42
         Top             =   3780
         Width           =   585
      End
      Begin VB.Frame fStatus 
         Caption         =   "Status"
         Height          =   795
         Left            =   120
         TabIndex        =   39
         Top             =   4710
         Width           =   3885
         Begin VB.Label lStatus 
            Caption         =   "Waiting for Action..."
            Height          =   375
            Left            =   150
            TabIndex        =   40
            Top             =   330
            Width           =   3585
         End
      End
      Begin VB.Frame fControl 
         Caption         =   "Control"
         Height          =   795
         Left            =   4230
         TabIndex        =   36
         Top             =   4710
         Width           =   2655
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H008080FF&
            Caption         =   "Cancel"
            Height          =   315
            Left            =   1350
            MaskColor       =   &H008080FF&
            TabIndex        =   38
            Top             =   300
            Width           =   1185
         End
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H0080FF80&
            Caption         =   "Update"
            Height          =   315
            Left            =   120
            MaskColor       =   &H0080FF80&
            TabIndex        =   37
            Top             =   300
            Width           =   1185
         End
      End
      Begin VB.TextBox txtExp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         TabIndex        =   35
         Text            =   "0"
         Top             =   870
         Width           =   975
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   3420
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   31
         Top             =   1800
         Width           =   510
      End
      Begin VB.TextBox txtMP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   390
         TabIndex        =   22
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   255
         Left            =   390
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtSpi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   420
         TabIndex        =   20
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtAgi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   255
         Left            =   420
         TabIndex        =   19
         Top             =   3180
         Width           =   735
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   420
         TabIndex        =   18
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Height          =   255
         Left            =   420
         TabIndex        =   17
         Top             =   2610
         Width           =   735
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   420
         TabIndex        =   16
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         TabIndex        =   2
         Text            =   "0"
         Top             =   570
         Width           =   975
      End
      Begin VB.Label lblPts 
         BackColor       =   &H000080FF&
         Caption         =   "Points:"
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   3810
         Width           =   495
      End
      Begin VB.Label labClass 
         Caption         =   "------"
         Height          =   165
         Left            =   3420
         TabIndex        =   34
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label labGender 
         Caption         =   "------"
         Height          =   165
         Left            =   3420
         TabIndex        =   33
         Top             =   930
         Width           =   915
      End
      Begin VB.Label labName 
         Caption         =   "------"
         Height          =   195
         Left            =   3420
         TabIndex        =   32
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label lblExamp 
         Caption         =   "(other chars from the same account)"
         Height          =   495
         Left            =   5310
         TabIndex        =   30
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite:"
         Height          =   225
         Left            =   2880
         TabIndex        =   29
         Top             =   2070
         Width           =   525
      End
      Begin VB.Label lblAccess 
         Caption         =   "Access:"
         Height          =   225
         Left            =   2820
         TabIndex        =   28
         Top             =   1470
         Width           =   525
      End
      Begin VB.Label lblClass 
         Caption         =   "Class:"
         Height          =   225
         Left            =   2940
         TabIndex        =   27
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label lblGender 
         Caption         =   "Gender:"
         Height          =   195
         Left            =   2790
         TabIndex        =   26
         Top             =   930
         Width           =   585
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   5220
         X2              =   6720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblSimilar 
         Caption         =   "Similar Characters"
         Height          =   165
         Left            =   5340
         TabIndex        =   25
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   165
         Left            =   2910
         TabIndex        =   24
         Top             =   690
         Width           =   495
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   2790
         X2              =   4560
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General"
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   390
         Width           =   1155
      End
      Begin VB.Line Line8 
         X1              =   90
         X2              =   1110
         Y1              =   3750
         Y2              =   3750
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   1080
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line6 
         X1              =   90
         X2              =   1080
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   1080
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line3 
         X1              =   390
         X2              =   390
         Y1              =   2400
         Y2              =   3780
      End
      Begin VB.Label lblSpi 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Spi:"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label lblAgi 
         BackColor       =   &H0080C0FF&
         Caption         =   "Agi:"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   3210
         Width           =   285
      End
      Begin VB.Label lblInt 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Int:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2940
         Width           =   315
      End
      Begin VB.Label lblEndu 
         BackColor       =   &H0080C0FF&
         Caption         =   "End:"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label lblStr 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Str:"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   2340
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   1080
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Label lblStats 
         BackColor       =   &H000080FF&
         Caption         =   "Stats"
         Height          =   165
         Left            =   90
         TabIndex        =   10
         Top             =   2130
         Width           =   1035
      End
      Begin VB.Label lblMP 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MP:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1770
         Width           =   255
      End
      Begin VB.Label lblHP 
         BackColor       =   &H0080FF80&
         Caption         =   "HP:"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   1470
         Width           =   405
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   1020
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label lblVitals 
         BackColor       =   &H0000FF00&
         Caption         =   "Vitals"
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblExp 
         Caption         =   "Exp: "
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   900
         Width           =   345
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.TextBox txtFake 
      Height          =   315
      Left            =   3210
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   450
      Width           =   495
   End
End
Attribute VB_Name = "frmCharEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public currentSprite As Long
Private textureNum As Long
Public maxSprites As Long
Private TmpIndex As Long

Private Sub cmbAccess_LostFocus()
    requestedPlayer.Access = cmbAccess.ListIndex
End Sub

Private Sub cmdCancel_Click()
    frmCharEditor.Width = 2865
    SendRequestAllCharacters
End Sub

Private Sub cmdFastCancel_Click()
   frmCharEditor.Width = 2865
   SendRequestAllCharacters
End Sub

Private Sub cmdFastUpdate_Click()
    frmCharEditor.Width = 2865
    SendCharacterUpdate
    SendRequestAllCharacters
End Sub

Private Sub cmdUpdate_Click()
    frmCharEditor.Width = 2865
    SendCharacterUpdate
    SendRequestAllCharacters
End Sub

Private Sub Form_Load()
    listCharacters.listItems.Clear
    listCharacters.ColumnHeaders.Item(1).Width = frmCharEditor.listCharacters.Width - 800
    LastCharSpriteTimer = timeGetTime
    frmCharEditor.Width = 2865
    upSprite.max = NumCharacters
    upSprite.min = 0
    txtFilter.MaxLength = NAME_LENGTH
End Sub

Public Sub ResetCharList()
    listCharacters.Sorted = False
    listCharacters.listItems.Clear
    Dim Length As Long, I As Long
    Length = UBound(charList)
    For I = 0 To Length
        listCharacters.listItems.Add , , charList(I, 0)
        listCharacters.listItems.Item(I + 1).SubItems(1) = charList(I, 1)
    Next
    listCharacters.ColumnHeaders.Item(2).Width = 800
End Sub

Public Sub fetchPlayerData()
    textureNum = -1
    txtLevel = requestedPlayer.Level
    txtEXP = requestedPlayer.exp
    txtHP.text = requestedPlayer.Vital(1)
    txtMP.text = requestedPlayer.Vital(2)
    txtStr.text = requestedPlayer.Stat(1)
    txtEnd.text = requestedPlayer.Stat(2)
    txtInt.text = requestedPlayer.Stat(3)
    txtAgi.text = requestedPlayer.Stat(4)
    txtSpi.text = requestedPlayer.Stat(5)
    txtPoints.text = requestedPlayer.Points
    
    txtSprite = requestedPlayer.Sprite
    
    labName.Caption = requestedPlayer.Name
    frmCharEditor.Width = 10035
    'Gender
    If requestedPlayer.Gender = 0 Then
        labGender.Caption = "Male"
    Else
        labGender.Caption = "Female"
    End If
    'Class
        labClass.Caption = Class(requestedPlayer.Class).Name
    'Access
    Select Case requestedPlayer.Access
        Case 0: cmbAccess.ListIndex = 0
        Case 1: cmbAccess.ListIndex = 1
        Case 2: cmbAccess.ListIndex = 2
        Case 3: cmbAccess.ListIndex = 3
        Case 4: cmbAccess.ListIndex = 4
        Case 5: cmbAccess.ListIndex = 5
    End Select
    SetSprite
End Sub

Private Sub SetSprite()
    If requestedPlayer.Sprite > 0 And requestedPlayer.Sprite <= NumCharacters Then
        If textureNum = -1 Then
            NumTextures = NumTextures + 1
            textureNum = NumTextures
            ReDim Preserve gTexture(textureNum)
        End If

        Tex_CharSprite.filepath = App.Path & "\data files\graphics\characters\" & Trim$(str$(requestedPlayer.Sprite)) & ".png"
        Tex_CharSprite.Texture = textureNum
        LoadTexture Tex_CharSprite
    Else
        picSprite.Picture = Nothing
    End If

End Sub

Private Sub frameCharPanel_Click()
    txtFake.SetFocus
End Sub

Private Sub listCharacters_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With listCharacters
        If .SortKey <> ColumnHeader.Index - 1 Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        Else
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
             Else
                 .SortOrder = lvwAscending
            End If
        End If
        .Sorted = -1
    End With
End Sub

Private Sub listCharacters_DblClick()
    ' RequestCharacterData - check if online and edit it
    If listCharacters.SelectedItem.text = "" Then Exit Sub
    frameCharPanel.Caption = "Editing - " & listCharacters.SelectedItem.text & " - " & listCharacters.SelectedItem.SubItems(1)
    SendRequestExtendedPlayerData (listCharacters.SelectedItem.text)
End Sub

Private Sub txtAgi_Change()
     correctValue txtAgi, requestedPlayer.Stat(4), 0, 65535
End Sub

Private Sub txtAgi_Click()
    selectValue txtAgi
End Sub

Private Sub txtAgi_GotFocus()
    selectValue txtAgi
End Sub

Private Sub txtAgi_LostFocus()
    reviseValue txtAgi, requestedPlayer.Stat(4)
End Sub

Private Sub txtEnd_Change()
     correctValue txtEnd, requestedPlayer.Stat(2), 0, 65535
End Sub

Private Sub txtEnd_Click()
    selectValue txtEnd
End Sub

Private Sub txtEnd_GotFocus()
    selectValue txtEnd
End Sub

Private Sub txtEnd_LostFocus()
    reviseValue txtEnd, requestedPlayer.Stat(2)
End Sub

Private Sub txtExp_Change()
     correctValue txtEXP, requestedPlayer.exp, 0, 9999999
End Sub

Private Sub txtExp_Click()
    selectValue txtEXP
End Sub

Private Sub txtEXP_GotFocus()
    selectValue txtEXP
End Sub

Private Sub txtExp_LostFocus()
    reviseValue txtEXP, requestedPlayer.exp
End Sub

Private Sub txtFilter_Change()
    Dim test As Long
    If txtFilter.text <> "" Then
        listCharacters.listItems.Clear
        
        Dim content As String, Length As Long, I As Long
        content = txtFilter.text
    
        Length = UBound(charList)
        listCharacters.Sorted = False
        For I = 0 To Length
            If InStr(LCase$(charList(I, 0)), LCase$(content)) <> 0 Then
                listCharacters.listItems.Add , , charList(I, 0)
                listCharacters.listItems.Item(listCharacters.listItems.Count).SubItems(1) = charList(I, 1)
            End If
        Next
        listCharacters.Sorted = True
    Else
        ResetCharList
    End If
End Sub

Private Sub txtHP_Change()
     correctValue txtHP, requestedPlayer.Vital(1), 0, 65535
End Sub

Private Sub txtHP_Click()
     selectValue txtHP
End Sub

Private Sub txtHP_GotFocus()
     selectValue txtHP
End Sub

Private Sub txtHP_LostFocus()
    reviseValue txtHP, requestedPlayer.Vital(1)
End Sub

Private Sub txtInt_Change()
     correctValue txtInt, requestedPlayer.Stat(3), 0, 65535
End Sub

Private Sub txtInt_Click()
     selectValue txtInt
End Sub

Private Sub txtInt_GotFocus()
     selectValue txtInt
End Sub

Private Sub txtInt_LostFocus()
    reviseValue txtInt, requestedPlayer.Stat(3)
End Sub

Private Sub txtLevel_Change()
     correctValue txtLevel, requestedPlayer.Level, 0, MAX_LEVEL
End Sub

Private Sub txtLevel_Click()
     selectValue txtLevel
End Sub

Private Sub txtLevel_GotFocus()
     selectValue txtLevel
End Sub

Private Sub txtLevel_LostFocus()
    reviseValue txtLevel, requestedPlayer.Level
End Sub

Private Sub txtMP_Change()
     correctValue txtMP, requestedPlayer.Vital(2), 0, 65535
End Sub

Private Sub txtMP_Click()
     selectValue txtMP
End Sub

Private Sub txtMP_GotFocus()
    selectValue txtMP
End Sub

Private Sub txtMP_LostFocus()
    reviseValue txtMP, requestedPlayer.Vital(2)
End Sub

Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub
Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 0) As Boolean
    Dim test As textBox
    
    If textBox.text = "" Then
        textBox.text = CStr(defaultVal)
        valueToChange = defaultVal
        correctValue = True
    End If

    If Len(textBox.text) = 1 And InStr(1, textBox.text, "-") = 1 Then
        correctValue = True
        Exit Function
    ElseIf Len(textBox.text) = 1 And IsNumeric(textBox.text) Then
        If verifyValue(textBox, min, max) Then
            valueToChange = textBox.text
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 0 And InStrRev(textBox.text, "-") = 0 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            valueToChange = textBox.text
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If

    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 1 And InStrRev(textBox.text, "-") = 1 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            valueToChange = textBox.text
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
        correctValue = False
        End If
        
    Else
        textBox.text = CStr(valueToChange)
        textBox.SelStart = Len(textBox.text)
        correctValue = False
    End If
End Function

Private Sub reviseValue(ByRef textBox As textBox, ByRef valueToChange)
    If Not IsNumeric(textBox.text) Then
        textBox.text = CStr(valueToChange)
    Else
        textBox.text = CStr(valueToChange)
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long)
    Dim Msg As String
    
    If (CLng(txtBox.text) >= min And CLng(txtBox.text) <= max) Then
        verifyValue = True
    Else
        Msg = " field accepts only values: " & CStr(min) & " < value < " & CStr(max) & "." & vbCrLf & "Reverting value..."
        verifyValue = False
    End If
End Function

Private Sub txtPoints_Change()
     correctValue txtPoints, requestedPlayer.Points, 0, 65535
End Sub

Private Sub txtPoints_Click()
     selectValue txtPoints
End Sub

Private Sub txtPoints_GotFocus()
    selectValue txtPoints
End Sub

Private Sub txtPoints_LostFocus()
    reviseValue txtPoints, requestedPlayer.Points
End Sub

Private Sub txtSpi_Change()
     correctValue txtSpi, requestedPlayer.Stat(5), 0, 65535
End Sub

Private Sub txtSpi_Click()
    selectValue txtSpi
End Sub

Private Sub txtSpi_GotFocus()
    selectValue txtSpi
End Sub

Private Sub txtSpi_LostFocus()
    reviseValue txtSpi, requestedPlayer.Stat(5)
End Sub

Private Sub txtSprite_Change()
     Dim OK As Boolean
     
     OK = correctValue(txtSprite, requestedPlayer.Sprite, 1, NumCharacters, 1)
     
     If OK Then
        SetSprite
     End If
End Sub

Private Sub txtSprite_Click()
    selectValue txtSprite
End Sub

Private Sub txtSprite_GotFocus()
    selectValue txtSprite
End Sub

Private Sub txtSprite_LostFocus()
    reviseValue txtSprite, requestedPlayer.Sprite
End Sub

Private Sub txtStr_Change()
     correctValue txtStr, requestedPlayer.Stat(1), 1, 65535
End Sub

Private Sub txtStr_Click()
    selectValue txtStr
End Sub

Private Sub txtStr_GotFocus()
    selectValue txtStr
End Sub

Private Sub txtStr_LostFocus()
    reviseValue txtStr, requestedPlayer.Stat(1)
End Sub

