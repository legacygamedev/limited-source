VERSION 5.00
Begin VB.Form frmInbox 
   Caption         =   "My Inbox"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075.188
   ScaleMode       =   0  'User
   ScaleWidth      =   7600.889
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstOutbox 
      Height          =   2460
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame frNew 
      Height          =   5775
      Left            =   3360
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtBody 
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtReceiver 
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frOutbox 
      Height          =   5775
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtReceiver2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtBody3 
         Height          =   4575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblSender 
         Caption         =   "To"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame frInbox 
      Height          =   5775
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtBody2 
         Height          =   4575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtSender 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOutboxSel 
      Caption         =   "Outbox"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdInboxSel 
      Caption         =   "Inbox"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward Message"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton cmdCloseInbox 
      Caption         =   "Close"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
   Begin VB.ListBox lstMail 
      Height          =   2460
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Message"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Message"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   3015
   End
End
Attribute VB_Name = "frmInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim x As Long
    x = MsgBox("Are you sure you want to discard your message?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
 
    frNew.Visible = False
    frmInbox.Width = 3450
    cmdForward.Enabled = True
    cmdOutboxSel.Enabled = True
    cmdInboxSel.Enabled = True
    cmdDelete.Enabled = True
    cmdCloseInbox.Enabled = True
    cmdNew.Enabled = True
    frInbox.Visible = False
    frOutbox.Visible = False
End Sub

Private Sub cmdCloseInbox_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim DltMsg As String
    Dim DltMsgNum As Long

    If frNew.Visible = True Then
        Exit Sub
    End If
    
    If lstMail.Visible = True Then
        DltMsg = lstMail.Text
        If DltMsg <> "" Then
            DltMsgNum = Mid$(DltMsg, 2, 4)
            Call RemoveMail(GetPlayerName(MyIndex), DltMsgNum, 1)
            lstMail.RemoveItem lstMail.ListIndex
        End If
    Else
            DltMsg = lstOutbox.Text
        If DltMsg <> "" Then
            DltMsgNum = Mid$(DltMsg, 2, 4)
            Call RemoveMail(GetPlayerName(MyIndex), DltMsgNum, 2)
            lstOutbox.RemoveItem lstOutbox.ListIndex
        End If
    End If
End Sub

Private Sub cmdForward_Click()
    If frNew.Visible = True Then
        Exit Sub
    End If
    Dim ListedMsg As Long
    Dim MsgID As String
    If lstMail.Visible = True Then
        MsgID = lstMail.Text
        If MsgID <> "" Then
            ListedMsg = Mid$(MsgID, 2, 4)
            Call GetMsgBody(GetPlayerName(MyIndex), ListedMsg, 1)
        End If
    Else
        MsgID = lstOutbox.Text
        If MsgID <> "" Then
            ListedMsg = Mid$(MsgID, 2, 4)
            Call GetMsgBody(GetPlayerName(MyIndex), ListedMsg, 2)
        End If
    End If
    frNew.Visible = True
    frmInbox.Width = 7875
    cmdForward.Enabled = False
    cmdOutboxSel.Enabled = False
    cmdInboxSel.Enabled = False
    cmdDelete.Enabled = False
    cmdCloseInbox.Enabled = False
    cmdNew.Enabled = False
    frInbox.Visible = False
    frOutbox.Visible = False
End Sub

Private Sub cmdInboxSel_Click()
    frOutbox.Visible = False
    frmInbox.Width = 3450
    lstOutbox.Visible = False
    lstMail.Visible = True
End Sub

Private Sub cmdNew_Click()
    If frNew.Visible = False Then
        txtBody.Text = vbNullString
        txtReceiver.Text = vbNullString
        txtSubject.Text = vbNullString
        frNew.Visible = True
        frmInbox.Width = 7875
        cmdForward.Enabled = False
        cmdOutboxSel.Enabled = False
        cmdInboxSel.Enabled = False
        cmdDelete.Enabled = False
        cmdCloseInbox.Enabled = False
        cmdNew.Enabled = False
        frInbox.Visible = False
        frOutbox.Visible = False
    End If
End Sub

Private Sub cmdOutboxSel_Click()
    frOutbox.Visible = False
    frmInbox.Width = 3450
    lstOutbox.Visible = True
    lstMail.Visible = False
End Sub

Private Sub cmdSend_Click()
        Call CheckChar(GetPlayerName(MyIndex), frmInbox.txtReceiver.Text, frmInbox.txtSubject.Text, frmInbox.txtBody.Text)
        Unload Me
End Sub

Private Sub Form_Load()
    Call CheckInbox(GetPlayerName(MyIndex), 1)
    Call CheckInbox(GetPlayerName(MyIndex), 2)
    Me.Width = 3450
End Sub

Private Sub lstMail_Click()
    If frNew.Visible = True Then
        Exit Sub
    End If
    
    frInbox.Visible = True
    frOutbox.Visible = False
    Dim ListedMsg As Long
    Dim MsgID As String
    MsgID = lstMail.Text
    If MsgID <> "" Then
        ListedMsg = Mid$(MsgID, 2, 4)
        Call GetMsgBody(GetPlayerName(MyIndex), ListedMsg, 1)
    End If
    frmInbox.Width = 7875
    frInbox.Visible = True
    frOutbox.Visible = False
End Sub

Private Sub lstOutbox_Click()
    If frNew.Visible = True Then
        Exit Sub
    End If
    
    frInbox.Visible = False
    frOutbox.Visible = True
    Dim ListedMsg As Long
    Dim MsgID As String
    MsgID = lstOutbox.Text
    If MsgID <> "" Then
        ListedMsg = Mid$(MsgID, 2, 4)
        Call GetMsgBody(GetPlayerName(MyIndex), ListedMsg, 2)
    End If
    frmInbox.Width = 7875
    frInbox.Visible = False
    frOutbox.Visible = True
End Sub
