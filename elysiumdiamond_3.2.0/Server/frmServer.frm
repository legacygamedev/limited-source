VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elysium Diamond Server"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Initialising..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   5295
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnureloadscript 
         Caption         =   "Reload scripts"
      End
      Begin VB.Menu mnurefreshaccesses 
         Caption         =   "Refresh accesses"
      End
      Begin VB.Menu mnushutdown 
         Caption         =   "Shutdown"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Show the pop-up menu
    If X = RightUp Then Me.PopupMenu mnu, 0, , , mnushutdown

End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, _
   ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_DataArrival(Index As Integer, _
   ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If
End Sub

'Private Sub tmrMasterTimer_Timer()
'Dim Time As Byte
'Dim Time2 As Byte
'Dim Time3 As Byte
'
'Time = Time + 1
'Time2 = Time2 + 1
'Time3 = Time3 + 1
'
'    Call ServerLogic
'
'    If Time = 2 Then
'
'        Call CheckSpawnMapItems
'
'        If tmrChatLogs = YES Then
'            Call LogChats
'        End If
'
'        Time = 0
'    End If
'
'    If Time3 = 10 Then
'
'        If PlayerTimer = YES Then
'            Call PlayerSaveTimer2
'        End If
'
'        Time3 = 0
'    End If
'
'    If Time2 = 120 Then
'
'        If tmrPlayerSave = YES Then
'            Call PlayerSaveTimer
'        End If
'
'        Time2 = 0
'    End If
'
'End Sub

Private Sub mnushutdown_Click()

    Call DestroyServer

End Sub

Private Sub mnurefreshaccesses_Click()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) = True Then
            If GetPlayerAccess(i) <> Val(GetVar(App.Path & "\Accounts\" & GetPlayerLogin(i) & ".ini", "CHAR" & Player(i).CharNum, "Access")) Then
                Call SetPlayerAccess(i, Val(GetVar(App.Path & "\Accounts\" & GetPlayerLogin(i) & ".ini", "CHAR" & Player(i).CharNum, "Access")))
                Call SendPlayerData(i)
                Call PlayerMsg(i, "The server has changed your access to " & GetPlayerAccess(i) & "!", White)
            End If
        End If
    Next i

    MsgBox "Accesses have been refreshed.", vbOKOnly, "Refreshed"
    
End Sub

Private Sub mnureloadscript_Click()

    If SCRIPTING = 1 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        Call AddLog("Scripts reloaded.", "serverlog.txt")
    End If

End Sub
