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

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Private Sub chkQuit_Click()
    If chkQuit.value = 0 Then
        chkRespond.value = 0
        optResponces(0).value = False
        optResponces(1).value = False
        optResponces(2).value = False
    End If
End Sub

Private Sub chkRespond_Click()
Dim i As Long
Dim Done As Boolean
Dim P As Long
Dim O As Long

    O = SpeechEditorCurrentNumber

    i = 1
    Do While Not Done And i < 4
        If optResponces(i - 1).value = True Then Done = True
        i = i + 1
    Loop
    
    If Not Done Then optResponces(0).value = True
    
    For P = 1 To 3
        If frmSpeech.chkRespond.value = 1 Then
            frmSpeech.optResponces(P - 1).Enabled = True
            frmSpeech.txtTalk(P - 1).Enabled = True
            frmSpeech.scrlGoTo(P - 1).Enabled = True
            frmSpeech.lblGoTo(P - 1).Enabled = True
            frmSpeech.chkExit(P - 1).Enabled = True
            
            If Speech(EditorIndex).num(O).Respond = P Then
                frmSpeech.optResponces(P - 1).value = True
            End If
        
            frmSpeech.txtTalk(P - 1).Text = Speech(EditorIndex).num(O).Responces(P).Text
            frmSpeech.scrlGoTo(P - 1).value = Speech(EditorIndex).num(O).Responces(P).GoTo
            frmSpeech.lblGoTo(P - 1).Caption = "Go to " & Speech(EditorIndex).num(O).Responces(P).GoTo
            frmSpeech.chkExit(P - 1).value = Speech(EditorIndex).num(O).Responces(P).Exit
        Else
            frmSpeech.optResponces(P - 1).Enabled = False
            frmSpeech.txtTalk(P - 1).Enabled = False
            frmSpeech.scrlGoTo(P - 1).Enabled = False
            frmSpeech.lblGoTo(P - 1).Enabled = False
            frmSpeech.chkExit(P - 1).Enabled = False
            
            frmSpeech.txtTalk(P - 1).Text = "Write a responce here."
            frmSpeech.scrlGoTo(P - 1).value = 0
            frmSpeech.lblGoTo(P - 1).Caption = "Go to 0"
            frmSpeech.chkExit(P - 1).value = 0
            
            optResponces(0).value = False
            optResponces(1).value = False
            optResponces(2).value = False
        End If
    Next P
End Sub

Private Sub chkScript_Click()
    If chkScript.value = 0 Then
        scrlScript.Visible = False
        scrlScript.value = 0
        lblScript.Visible = False
    Else
        scrlScript.Visible = True
        scrlScript.value = Speech(EditorIndex).num(SpeechEditorCurrentNumber).Script
        lblScript.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Call SendData("NEEDSPEECH" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    
    InSpeechEditor = False
    Unload frmSpeech
End Sub

Private Sub cmdDone_Click()
Dim O As Long

    O = SpeechEditorCurrentNumber
    
    ' Save it
    
    Speech(EditorIndex).num(O).Exit = chkQuit.value
    Speech(EditorIndex).num(O).Text = txtMainTalk.Text
    Speech(EditorIndex).num(O).Script = scrlScript.value
    If optSaid(0).value = True Then
        Speech(EditorIndex).num(O).SaidBy = 0
    Else
        Speech(EditorIndex).num(O).SaidBy = 1
    End If
    
    For P = 1 To 3
        If optResponces(P - 1).value = True Then Speech(EditorIndex).num(O).Respond = P
        If chkRespond = 0 Then Speech(EditorIndex).num(O).Respond = 0
        
        Speech(EditorIndex).num(O).Responces(P).Exit = chkExit(P - 1).value
        Speech(EditorIndex).num(O).Responces(P).GoTo = scrlGoTo(P - 1).value
        Speech(EditorIndex).num(O).Responces(P).Text = txtTalk(P - 1).Text
    Next P
    
    Speech(EditorIndex).name = txtName.Text
    
    Call SendSaveSpeech(EditorIndex)
    
    InSpeechEditor = False
    Unload frmSpeech
End Sub

Private Sub optResponces_Click(Index As Integer)
    If chkRespond.value = 0 Then chkRespond.value = 1
End Sub

Private Sub scrlGoTo_Change(Index As Integer)
    lblGoTo(Index).Caption = "Go to " & scrlGoTo(Index).value
End Sub

Private Sub scrlNumber_Change()
Dim i As Long
Dim O As Long
Dim P As Long

    i = scrlNumber.value
    O = SpeechEditorCurrentNumber
    
    ' Save it
    Speech(EditorIndex).num(O).Exit = chkQuit.value
    Speech(EditorIndex).num(O).Text = txtMainTalk.Text
    Speech(EditorIndex).num(O).Script = scrlScript.value
    If optSaid(0).value = True Then
        Speech(EditorIndex).num(O).SaidBy = 0
    Else
        Speech(EditorIndex).num(O).SaidBy = 1
    End If
    
    For P = 1 To 3
        If optResponces(P - 1).value = True Then Speech(EditorIndex).num(O).Respond = P
        If chkRespond = 0 Then Speech(EditorIndex).num(O).Respond = 0
        
        Speech(EditorIndex).num(O).Responces(P).Exit = chkExit(P - 1).value
        Speech(EditorIndex).num(O).Responces(P).GoTo = scrlGoTo(P - 1).value
        Speech(EditorIndex).num(O).Responces(P).Text = txtTalk(P - 1).Text
    Next P
    
    ' Load new stuff
    chkQuit.value = Speech(EditorIndex).num(i).Exit
    txtMainTalk.Text = Speech(EditorIndex).num(i).Text
    optSaid(Speech(EditorIndex).num(i).SaidBy).value = True
    
    If Speech(EditorIndex).num(i).Respond > 0 Then
        chkRespond.value = 1
    Else
        chkRespond.value = 0
    End If
    
    If Speech(EditorIndex).num(i).Script = 0 Then
        chkScript.value = 0
        scrlScript.value = 0
        lblScript.Caption = "0"
        scrlScript.Visible = False
        lblScript.Visible = False
    Else
        chkScript.value = 1
        scrlScript.value = Speech(EditorIndex).num(i).Script
        lblScript.Caption = Speech(EditorIndex).num(i).Script
        scrlScript.Visible = True
        lblScript.Visible = True
    End If

    For P = 1 To 3
        If chkRespond.value = 1 Then
            optResponces(P - 1).Enabled = True
            txtTalk(P - 1).Enabled = True
            scrlGoTo(P - 1).Enabled = True
            lblGoTo(P - 1).Enabled = True
            chkExit(P - 1).Enabled = True
            
            If Speech(EditorIndex).num(i).Respond = P Then
                optResponces(P - 1).value = True
            End If
        
            txtTalk(P - 1).Text = Speech(EditorIndex).num(i).Responces(P).Text
            scrlGoTo(P - 1).value = Speech(EditorIndex).num(i).Responces(P).GoTo
            lblGoTo(P - 1).Caption = "Go to " & Speech(EditorIndex).num(i).Responces(P).GoTo
            chkExit(P - 1).value = Speech(EditorIndex).num(i).Responces(P).Exit
        Else
            optResponces(P - 1).Enabled = False
            txtTalk(P - 1).Enabled = False
            scrlGoTo(P - 1).Enabled = False
            lblGoTo(P - 1).Enabled = False
            chkExit(P - 1).Enabled = False
        End If
    Next P
    
    If scrlNumber.value = 0 Then
        chkQuit.Enabled = False
        chkScript.Enabled = False
    Else
        chkQuit.Enabled = True
        chkScript.Enabled = True
    End If
    
    SpeechEditorCurrentNumber = i
    
    lblSection.Caption = i
End Sub

Private Sub scrlScript_Change()
    lblScript.Caption = scrlScript.value
End Sub

Private Sub txtName_Change()
    If Trim(txtName.Text) = "" Then
        lblWarn.Visible = True
    Else
        lblWarn.Visible = False
    End If
End Sub
