VERSION 5.00
Begin VB.Form frmMapReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Report"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmMapReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   2220
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpenMaps 
      Caption         =   "Open Maps"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   1185
   End
   Begin VB.CommandButton cmdWarp 
      Caption         =   "Warp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   765
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maps"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.ListBox lstMaps 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmMapReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_MAPS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_MAPS Then Exit Sub
    
    If Res <= GetPlayerMap(MyIndex) Then
        Call AlertMsg("You can't resize to a map number that would affect the map you are currently on!")
        Exit Sub
    End If
    
    Call SendChangeDataSize(Res, EDITOR_MAP)
    
    MAX_MAPS = Res
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Unload frmMapReport
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdOpenMaps_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then Exit Sub
    SendOpenMaps
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdOpenMaps_Click", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdWarp_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then Exit Sub
    Call WarpTo(lstMaps.ListIndex + 1)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdWarp_Click", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstMaps_DblClick()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then Exit Sub
    Call WarpTo(lstMaps.ListIndex + 1)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstMaps_DblClick", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 0 To lstMaps.ListCount - 1
        Find = Trim$(I + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstMaps.List(I)) >= Len(Find) Then
            If UCase$(Mid$(Trim$(lstMaps.List(I)), 1, Len(Find))) = UCase$(Find) Then
                lstMaps.ListIndex = I
                Exit For
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_Change", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmMapReport", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
