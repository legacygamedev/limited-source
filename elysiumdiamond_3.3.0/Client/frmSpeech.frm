VERSION 5.00
Begin VB.Form frmSpeech 
   Caption         =   "Speech Editor"
   ClientHeight    =   6150
   ClientLeft      =   1725
   ClientTop       =   1680
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8040
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Save"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame frameNumber 
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7815
      Begin VB.HScrollBar scrlScript 
         Height          =   255
         Left            =   2400
         Max             =   100
         TabIndex        =   31
         Top             =   240
         Width           =   4695
      End
      Begin VB.CheckBox chkScript 
         Caption         =   "Script?"
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
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkQuit 
         Caption         =   "Exit?"
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkRespond 
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   2040
         Width           =   255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Responces "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   7575
         Begin VB.CheckBox chkExit 
            Caption         =   "Exit?"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   24
            Top             =   1560
            Width           =   2295
         End
         Begin VB.HScrollBar scrlGoTo 
            Height          =   255
            Index           =   2
            Left            =   5160
            Max             =   10
            TabIndex        =   23
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optResponces 
            Caption         =   "Option2"
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtTalk 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   5160
            TabIndex        =   21
            Text            =   "Write a responce here."
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkExit 
            Caption         =   "Exit?"
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
            Index           =   1
            Left            =   2640
            TabIndex        =   19
            Top             =   1560
            Width           =   2295
         End
         Begin VB.HScrollBar scrlGoTo 
            Height          =   255
            Index           =   1
            Left            =   2640
            Max             =   10
            TabIndex        =   18
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optResponces 
            Caption         =   "Option2"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtTalk 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   16
            Text            =   "Write a responce here."
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkExit 
            Caption         =   "Exit?"
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
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   2295
         End
         Begin VB.HScrollBar scrlGoTo 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   10
            TabIndex        =   12
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtTalk 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Text            =   "Write a responce here."
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton optResponces 
            Caption         =   "Option2"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblGoTo 
            Caption         =   "Go To 0"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   25
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblGoTo 
            Caption         =   "Go to 0"
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
            Index           =   1
            Left            =   2640
            TabIndex        =   20
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblGoTo 
            Caption         =   "Go to 0"
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
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.OptionButton optSaid 
         Caption         =   "Player"
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
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton optSaid 
         Caption         =   "NPC"
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
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMainTalk 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmSpeech.frx":0000
         Top             =   600
         Width           =   7335
      End
      Begin VB.Label lblScript 
         Caption         =   "0"
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
         Left            =   7200
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Said by:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Responable?"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.HScrollBar scrlNumber 
      Height          =   255
      Left            =   240
      Max             =   10
      TabIndex        =   0
      Top             =   840
      Width           =   7575
   End
   Begin VB.Label lblWarn 
      Caption         =   "You MUST have a name to save the speech."
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
      Left            =   720
      TabIndex        =   33
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
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
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      TabIndex        =   29
      Top             =   600
      Width           =   7815
   End
End
Attribute VB_Name = "frmSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Private Sub chkQuit_Click()
    If chkQuit.Value = 0 Then
        chkRespond.Value = 0
        optResponces(0).Value = False
        optResponces(1).Value = False
        optResponces(2).Value = False
    End If
End Sub

Private Sub chkRespond_Click()
Dim I As Long
Dim done As Boolean
Dim p As Long
Dim O As Long

    O = SpeechEditorCurrentNumber

    I = 1
    Do While Not done And I < 4
        If optResponces(I - 1).Value = True Then done = True
        I = I + 1
    Loop
    
    If Not done Then optResponces(0).Value = True
    
    For p = 1 To 3
        If frmSpeech.chkRespond.Value = 1 Then
            frmSpeech.optResponces(p - 1).Enabled = True
            frmSpeech.txtTalk(p - 1).Enabled = True
            frmSpeech.scrlGoTo(p - 1).Enabled = True
            frmSpeech.lblGoTo(p - 1).Enabled = True
            frmSpeech.chkExit(p - 1).Enabled = True
            
            If Speech(EditorIndex).Num(O).Respond = p Then
                frmSpeech.optResponces(p - 1).Value = True
            End If
        
            frmSpeech.txtTalk(p - 1).Text = Speech(EditorIndex).Num(O).Responces(p).Text
            frmSpeech.scrlGoTo(p - 1).Value = Speech(EditorIndex).Num(O).Responces(p).GoTo
            frmSpeech.lblGoTo(p - 1).Caption = "Go to " & Speech(EditorIndex).Num(O).Responces(p).GoTo
            frmSpeech.chkExit(p - 1).Value = Speech(EditorIndex).Num(O).Responces(p).Exit
        Else
            frmSpeech.optResponces(p - 1).Enabled = False
            frmSpeech.txtTalk(p - 1).Enabled = False
            frmSpeech.scrlGoTo(p - 1).Enabled = False
            frmSpeech.lblGoTo(p - 1).Enabled = False
            frmSpeech.chkExit(p - 1).Enabled = False
            
            frmSpeech.txtTalk(p - 1).Text = "Write a responce here."
            frmSpeech.scrlGoTo(p - 1).Value = 0
            frmSpeech.lblGoTo(p - 1).Caption = "Go to 0"
            frmSpeech.chkExit(p - 1).Value = 0
            
            optResponces(0).Value = False
            optResponces(1).Value = False
            optResponces(2).Value = False
        End If
    Next p
End Sub

Private Sub chkScript_Click()
    If chkScript.Value = 0 Then
        scrlScript.Visible = False
        scrlScript.Value = 0
        lblScript.Visible = False
    Else
        scrlScript.Visible = True
        scrlScript.Value = Speech(EditorIndex).Num(SpeechEditorCurrentNumber).Script
        lblScript.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Call SendData(NEEDSPEECH_CHAR & SEP_CHAR & EditorIndex & END_CHAR)
    
    InSpeechEditor = False
    Unload frmSpeech
End Sub

Private Sub cmdDone_Click()
Dim O As Long

    O = SpeechEditorCurrentNumber
    
    ' Save it
    
    Speech(EditorIndex).Num(O).Exit = chkQuit.Value
    Speech(EditorIndex).Num(O).Text = txtMainTalk.Text
    Speech(EditorIndex).Num(O).Script = scrlScript.Value
    If optSaid(0).Value = True Then
        Speech(EditorIndex).Num(O).SaidBy = 0
    Else
        Speech(EditorIndex).Num(O).SaidBy = 1
    End If
    
    For p = 1 To 3
        If optResponces(p - 1).Value = True Then Speech(EditorIndex).Num(O).Respond = p
        If chkRespond = 0 Then Speech(EditorIndex).Num(O).Respond = 0
        
        Speech(EditorIndex).Num(O).Responces(p).Exit = chkExit(p - 1).Value
        Speech(EditorIndex).Num(O).Responces(p).GoTo = scrlGoTo(p - 1).Value
        Speech(EditorIndex).Num(O).Responces(p).Text = txtTalk(p - 1).Text
    Next p
    
    Speech(EditorIndex).Name = txtName.Text
    
    Call SendSaveSpeech(EditorIndex)
    
    InSpeechEditor = False
    Unload frmSpeech
End Sub

Private Sub optResponces_Click(Index As Integer)
    If chkRespond.Value = 0 Then chkRespond.Value = 1
End Sub

Private Sub scrlGoTo_Change(Index As Integer)
    lblGoTo(Index).Caption = "Go to " & scrlGoTo(Index).Value
End Sub

Private Sub scrlNumber_Change()
Dim I As Long
Dim O As Long
Dim p As Long

    I = scrlNumber.Value
    O = SpeechEditorCurrentNumber
    
    ' Save it
    Speech(EditorIndex).Num(O).Exit = chkQuit.Value
    Speech(EditorIndex).Num(O).Text = txtMainTalk.Text
    Speech(EditorIndex).Num(O).Script = scrlScript.Value
    If optSaid(0).Value = True Then
        Speech(EditorIndex).Num(O).SaidBy = 0
    Else
        Speech(EditorIndex).Num(O).SaidBy = 1
    End If
    
    For p = 1 To 3
        If optResponces(p - 1).Value = True Then Speech(EditorIndex).Num(O).Respond = p
        If chkRespond = 0 Then Speech(EditorIndex).Num(O).Respond = 0
        
        Speech(EditorIndex).Num(O).Responces(p).Exit = chkExit(p - 1).Value
        Speech(EditorIndex).Num(O).Responces(p).GoTo = scrlGoTo(p - 1).Value
        Speech(EditorIndex).Num(O).Responces(p).Text = txtTalk(p - 1).Text
    Next p
    
    ' Load new stuff
    chkQuit.Value = Speech(EditorIndex).Num(I).Exit
    txtMainTalk.Text = Speech(EditorIndex).Num(I).Text
    optSaid(Speech(EditorIndex).Num(I).SaidBy).Value = True
    
    If Speech(EditorIndex).Num(I).Respond > 0 Then
        chkRespond.Value = 1
    Else
        chkRespond.Value = 0
    End If
    
    If Speech(EditorIndex).Num(I).Script = 0 Then
        chkScript.Value = 0
        scrlScript.Value = 0
        lblScript.Caption = "0"
        scrlScript.Visible = False
        lblScript.Visible = False
    Else
        chkScript.Value = 1
        scrlScript.Value = Speech(EditorIndex).Num(I).Script
        lblScript.Caption = Speech(EditorIndex).Num(I).Script
        scrlScript.Visible = True
        lblScript.Visible = True
    End If

    For p = 1 To 3
        If chkRespond.Value = 1 Then
            optResponces(p - 1).Enabled = True
            txtTalk(p - 1).Enabled = True
            scrlGoTo(p - 1).Enabled = True
            lblGoTo(p - 1).Enabled = True
            chkExit(p - 1).Enabled = True
            
            If Speech(EditorIndex).Num(I).Respond = p Then
                optResponces(p - 1).Value = True
            End If
        
            txtTalk(p - 1).Text = Speech(EditorIndex).Num(I).Responces(p).Text
            scrlGoTo(p - 1).Value = Speech(EditorIndex).Num(I).Responces(p).GoTo
            lblGoTo(p - 1).Caption = "Go to " & Speech(EditorIndex).Num(I).Responces(p).GoTo
            chkExit(p - 1).Value = Speech(EditorIndex).Num(I).Responces(p).Exit
        Else
            optResponces(p - 1).Enabled = False
            txtTalk(p - 1).Enabled = False
            scrlGoTo(p - 1).Enabled = False
            lblGoTo(p - 1).Enabled = False
            chkExit(p - 1).Enabled = False
        End If
    Next p
    
    If scrlNumber.Value = 0 Then
        chkQuit.Enabled = False
        chkScript.Enabled = False
    Else
        chkQuit.Enabled = True
        chkScript.Enabled = True
    End If
    
    SpeechEditorCurrentNumber = I
    
    lblSection.Caption = I
End Sub

Private Sub scrlScript_Change()
    lblScript.Caption = scrlScript.Value
End Sub

Private Sub txtName_Change()
    If Trim$(txtName.Text) = vbNullString Then
        lblWarn.Visible = True
    Else
        lblWarn.Visible = False
    End If
End Sub
