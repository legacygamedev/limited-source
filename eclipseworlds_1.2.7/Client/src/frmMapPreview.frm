VERSION 5.00
Begin VB.Form frmMapPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Preview - Map #"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   Icon            =   "frmMapPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   0
      ScaleHeight     =   2520
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frmMapPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents subclasser As cSelfSubHookCallback
Attribute subclasser.VB_VarHelpID = -1
Private Const WM_MOVE       As Long = &H3
Public redrawMapPreview As Boolean

Private Function scaleBase() As Byte
    If Map.MaxX >= Map.MaxY Then
        scaleBase = CByte(255 \ Map.MaxX)
     Else
        scaleBase = CByte(255 \ Map.MaxY)
    End If
End Function

Public Sub RecalcuateDimensions()
    Dim mapScale As Byte
    frmMapPreview.Caption = "Map Preview - #" & Player(MyIndex).Map
    mapScale = scaleBase()
    Width = PixelsToTwips(Map.MaxX * mapScale, 0)
    Height = PixelsToTwips(Map.MaxY * mapScale, 1)
    picMapPreview.Width = Map.MaxX * mapScale
    picMapPreview.Height = Map.MaxY * mapScale
    Move frmMain.Left - Width, frmMain.Top
    redrawMapPreview = True
End Sub

Private Sub Form_Load()
    If subclasser Is Nothing Then
        Set subclasser = New cSelfSubHookCallback
    End If
    
    If subclasser.ssc_Subclass(Me.hWnd, ByVal 1, 1, Me) Then
        subclasser.ssc_AddMsg Me.hWnd, eMsgWhen.MSG_BEFORE, WM_MOVE
    End If
    
    RecalcuateDimensions
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mapPreviewSwitch.Value = 0
    subclasser.ssc_UnSubclass Me.hWnd
    Set subclasser = Nothing
End Sub

Private Sub picMapPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GetPlayerAccess(MyIndex) >= STAFF_MAPPER Then
        If Button = vbLeftButton Then
            AdminWarp CLng(TwipsToPixels(X, 0) \ scaleBase), CLng(TwipsToPixels(Y, 1) \ scaleBase)
        End If
    End If
End Sub

Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
    Select Case uMsg
        Case WM_MOVE
            frmMain.Move frmMapPreview.Left + frmMapPreview.Width, frmMapPreview.Top
            frmEditor_Map.Move frmMain.Left - frmEditor_Map.Width - 136, frmMain.Top + frmMapPreview.Height
            bHandled = 1
    End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
End Sub
