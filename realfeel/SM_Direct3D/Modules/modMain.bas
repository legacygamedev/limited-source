Attribute VB_Name = "modMain"
Option Explicit

Public SM3D As clsDirect3D
Public SM_RENDER As Boolean
Public Pos_X As Single
Public Pos_Y As Single

'//These aren't really required - they'll just show us what the frame rate is...
Public Declare Function GetTickCount Lib "kernel32" () As Long '//This is used to get the frame rate.
Public LastTimeCheckFPS As Long '//When did we last check the frame rate?
Public FramesDrawn As Long '//How many frames have been drawn
Public FrameRate As Long '//What the current frame rate is.....

Sub Main()
On Error Resume Next
frmMain.Show

Set SM3D = New clsDirect3D

Pos_X = 10
Pos_Y = 10

SM3D.SetTriStripTopLeft Pos_X, Pos_Y, 0, 0, 32, 32
SM_RENDER = True
SM3D.SetRenderTex
SM3D.SetTriStripTopLeft Pos_X, Pos_Y, 0, 0, 32, 32, "&HFFFFFFFF"
Do While SM_RENDER
    
    SM3D.Render
    'DoEvents
    
    ' Make it move around
    'If Pos_X < 100 And Pos_Y = 10 Then
    '    Pos_X = Pos_X + 10
    'ElseIf Pos_X = 100 And Pos_Y < 100 Then
    '    Pos_Y = Pos_Y + 10
    'ElseIf Pos_X > 10 And Pos_Y = 100 Then
    '    Pos_X = Pos_X - 10
    'ElseIf Pos_X = 10 And Pos_Y > 10 Then
    '    Pos_Y = Pos_Y - 10
    'End If
    
    'Calculate the frame rate; how this is done isn't greatly important
    'So dont worry about understanding it yet...
    If GetTickCount - LastTimeCheckFPS >= 1000 Then
        LastTimeCheckFPS = GetTickCount
        FrameRate = FramesDrawn '//Store the frame count
        FramesDrawn = 0 '//Reset the counter
        frmMain.Caption = "SM Direct3D {" & FrameRate & "fps}" '//Display it on screen
    Else
        FramesDrawn = FramesDrawn + 1
    End If
    DoEvents
Loop
End Sub
