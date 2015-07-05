Attribute VB_Name = "modDirectX8"
Option Explicit
'Added DirectX8 module (04/23/07)
'Moved DirectMusic and DirectSound to this module (04/23/07)
'Created Direct3D processes here
'-smchronos

Public Function InitDirectX8() As Boolean
InitDirectX8 = False
'Let's start up that DirectX8 instance
Set DX8 = New DirectX8

Set Direct3D = New clsDirect3D
Set DirectMusic = New clsDirectMusic
Set DirectSound = New clsDirectSound

'Load Direct3D
Call Direct3D.InitDirect3D(frmDualSolace.picScreen.hwnd, frmDualSolace.picScreen.ScaleWidth, frmDualSolace.picScreen.ScaleHeight, False, True)

'load DirectMusic
Call DirectMusic.InitDirectMusic

'load DirectSound (not tested)
'Call DirectSound.InitDirectSound

'load pictureboxes for the editor, they are actually used for blitting right now
Call SetStatus("Loading graphics for equipment and items...")
DoEvents
Call SetPicSize(App.Path + GFX_PATH + "items" + GFX_EXT, frmEditor.picItems)
frmEditor.picItems.Picture = LoadPicture(App.Path & GFX_PATH & "items" & GFX_EXT)

'load these with the standard setting
AllowMovement = False
AttributeDisplay = True
DepictAttributeTiles = True

InitDirectX8 = True
End Function

Sub DestroyDirectX8()
    Call Direct3D.UnloadDirect3D(True)
    Call DirectMusic.DestroyDirectMusic
    Call DirectSound.DestroyDirectSound
    Set DX8 = Nothing
End Sub

