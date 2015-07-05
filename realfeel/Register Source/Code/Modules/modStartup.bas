Attribute VB_Name = "modStartup"
Option Explicit

Public Sub Main()
' Check to see if the configuration ini exists
' If not, make it.
If Not Exists(App.Path & "\config.ini") Then Call MakeConfigINI
If Not Exists(App.Path & "\exts.ini") Then Call MakeExtensionINI
If Dir(App.Path & "\Extensions\", vbDirectory) = "" Then Call MkDir(App.Path & "\Extensions\")

' Get path
Ext_Path = Read(App.Path & "\config.ini", "Path ")

' Run extension check
Call CheckReg(Ext_Path)

' Make form visible
frmRegFiles.Visible = True

If frmRegFiles.lstRegFiles.ListCount <> 0 Then
    frmRegFiles.lstRegFiles.ListIndex = 0
    Call frmRegFiles.lstRegFiles_Click
Else
    frmRegFiles.lblRegName.Caption = "---"
    frmRegFiles.lblRegDesc.Caption = "No file found!"
    frmRegFiles.lblRegFileStatus.Caption = "---"
End If
End Sub

Public Sub MakeConfigINI()
Dim f As Integer
f = FreeFile

Open (App.Path & "\config.ini") For Output As #f
    Print #f, ("--Dual Solace Registering Application--")
    Print #f, ("--Version: " & VERSION & "--")
    Print #f, ("")
    Print #f, ("Compiler: Visual Basic 6")
    Print #f, ("Created by smchronos")
    Print #f, ("")
    Print #f, ("Certified by Dual Solace 2008")
    Print #f, ""
    Print #f, "Path " & App.Path & "\Extensions\"
    'Print #f, "Do not delete the line below!"
    'Print #f, "[ENDFILE]"
Close #f
End Sub

Public Sub MakeExtensionINI()
Dim f As Integer
f = FreeFile

Open (App.Path & "\exts.ini") For Output As #f
    Print #f, ("--Dual Solace Registering Application--")
    Print #f, ("--Version: " & VERSION & "--")
    Print #f, ("")
    Print #f, ("Compiler: Visual Basic 6")
    Print #f, ("Created by smchronos")
    Print #f, ("")
    Print #f, ("Certified by Dual Solace 2008")
    Print #f, ""
    Print #f, "Information regarding extensions:"
    Print #f, "DX8VB.DLL=The Microsoft DirectX8 dynamic library. Used to display graphics, play music, and run sound."
    Print #f, "MSCOMM32.OCX=A Microsoft control containing common controls used in Visual Basic 6.0 applications."
    Print #f, "MSSTDFMT.DLL=A Microsoft dynamic library containing necessary Visual Studio information."
    Print #f, "MSWINSCK.OCX=The Microsoft Winsock control. Used to transmit data between client and server programs."
    Print #f, "RICHTX32.OCX=The Microsoft Rich Text Box control. Enables Rich Text Boxes."
    Print #f, "TABCTL32.OCX=A Microsoft control."
    'Print #f, "Do not delete the line below!"
    'Print #f, "[ENDFILE]"
Close #f
End Sub
