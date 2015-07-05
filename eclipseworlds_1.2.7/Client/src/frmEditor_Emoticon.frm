VERSION 5.00
Begin VB.Form frmEditor_Emoticon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   Icon            =   "frmEditor_Emoticon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3900
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6420
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   2055
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   3975
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3240
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   540
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picEmoticon 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   128
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Width           =   1920
            End
         End
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   120
         Max             =   1000
         TabIndex        =   3
         Top             =   1680
         Value           =   1
         Width           =   3735
      End
      Begin VB.TextBox txtCommand 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Command:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblEmoticon 
         Caption         =   "Emoticon: 0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Emoticon List"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   1230
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmEditor_Emoticon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpIndex As Long

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_EMOTICONS
        If Emoticon_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_EMOTICONS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_EMOTICONS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_EMOTICON)
    
    Unload frmEditor_Emoticon
    MAX_EMOTICONS = Res
    ReDim Emoticon(MAX_EMOTICONS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_EMOTICON).FontBold = False
    frmAdmin.picEye(EDITOR_EMOTICON).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Emoticon
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearEmoticon EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Emoticon(EditorIndex).Command, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    EmoticonEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    Dim I As Long, n As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_EMOTICONS
        ' Loop through a second time to compare if any match
        For n = 1 To MAX_EMOTICONS
            If Not Trim$(Emoticon(I).Command) = "/" And Not Trim$(Emoticon(n).Command) = "/" Then
                ' Make sure they are not the same one
                If Not I = n Then
                    If Trim$(Emoticon(I).Command) = Trim$(Emoticon(n).Command) Then
                        AlertMsg "There is more than one command that uses " & Trim$(txtCommand.text) & "!", True
                        Exit Sub
                    End If
                End If
            End If
        Next
    Next
    
    EditorSave = True
    Call EmoticonEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub


Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EmoticonEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "1stIndex_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlEmoticon_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblEmoticon.Caption = "Emoticon: " & scrlEmoticon.Value
    Emoticon(EditorIndex).Pic = scrlEmoticon.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlEmoticon_Change", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCommand_Validate(Cancel As Boolean)
    Dim I As Long, TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Make sure we have a slash
    If Not Left$(txtCommand.text, 1) = "/" Then
        If Trim$(txtCommand.text) = vbNullString Then
            txtCommand.text = "/"
        Else
            txtCommand.text = "/" & txtCommand.text
            txtCommand.SelStart = Len(txtCommand.text)
        End If
    End If
    
    TmpIndex = lstIndex.ListIndex
    Emoticon(EditorIndex).Command = Trim$(txtCommand.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Emoticon(EditorIndex).Command, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtCommand_Validate", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_EMOTICONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.UnsubDaFocus Me.hWnd
    
    If EditorSave = False Then
        EmoticonEditorCancel
    Else
        EditorSave = False
    End If
    
    frmAdmin.chkEditor(EDITOR_EMOTICON).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.SubDaFocus Me.hWnd
    
    ' Set max values
    txtSearch.MaxLength = NAME_LENGTH
    txtCommand.MaxLength = NAME_LENGTH
    scrlEmoticon.max = NumEmoticons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCommand_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtCommand.SelStart = Len(txtCommand)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtcommand_GotFocus", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 0 To lstIndex.ListCount - 1
        Find = Trim$(I + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstIndex.List(I)) >= Len(Find) Then
            If UCase$(Mid$(Trim$(lstIndex.List(I)), 1, Len(Find))) = UCase$(Find) Then
                lstIndex.ListIndex = I
                Exit For
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_Change", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        cmdClose_Click
        KeyAscii = 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Emoticon(EditorIndex)), ByVal VarPtr(Emoticon(TmpIndex + 1)), LenB(Emoticon(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Emoticon(EditorIndex).Command), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Emoticon", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
