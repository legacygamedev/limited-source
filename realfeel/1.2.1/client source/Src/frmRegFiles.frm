VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRegFiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RealFeel Registering Service"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip Register"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lstRegFiles 
      Height          =   1425
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar pbRegProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRegDesc 
      Caption         =   "REG DESCRIPTION"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblRegName 
      Alignment       =   2  'Center
      Caption         =   "NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblRegStatus 
      Alignment       =   2  'Center
      Caption         =   "STATUS"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4335
   End
End
Attribute VB_Name = "frmRegFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRegister_Click()
Dim n As Byte
Dim FileName As String, FileExt As String

' Set the status
frmRegFiles.lblRegStatus.Caption = "Preparing setup..."
frmRegFiles.pbRegProgress.Value = 0
frmRegFiles.pbRegProgress.Max = frmRegFiles.lstRegFiles.ListCount

' Drop the form down
frmRegFiles.Height = 3450

Debug.Print DLL_PATH

For n = 0 To frmRegFiles.lstRegFiles.ListCount - 1
        FileName = Left$(frmRegFiles.lstRegFiles.List(n), Len(frmRegFiles.lstRegFiles.List(n)) - 4)
        FileExt = Right$(frmRegFiles.lstRegFiles.List(n), 3)
        frmRegFiles.lblRegStatus.Caption = "Creating " & FileName & " batch file..."
        DoEvents
        Call MakeRegBatch(FileName, FileExt, App.Path & DLL_PATH)
        frmRegFiles.lblRegStatus.Caption = "Executing " & FileName & " batch file..."
        DoEvents
        Call RunRegBatch(FileName, App.Path & DLL_PATH, True)
        frmRegFiles.lblRegStatus.Caption = "Deleting " & FileName & " batch file..."
        DoEvents
    frmRegFiles.lblRegStatus.Caption = "Preparing next batch file..."
    frmRegFiles.pbRegProgress.Value = n + 1
    DoEvents
Next n

Call MsgBox("The system reads that all files have been successfully registered! If any problems are encountered, please report it to www.dualsolace.com/forum!")

' Now continue setup
frmRegFiles.Visible = False
Call Main2
Unload Me
End Sub

Private Sub cmdSkip_Click()
frmRegFiles.Visible = False
Call Main2
Unload Me
End Sub

Private Sub Form_Load()
    frmRegFiles.lstRegFiles.ListIndex = 0
    Call lstRegFiles_Click
End Sub

Private Sub lstRegFiles_Click()
' Load up the information
frmRegFiles.lblRegName.Caption = frmRegFiles.lstRegFiles.List(frmRegFiles.lstRegFiles.ListIndex)
Select Case UCase$(frmRegFiles.lstRegFiles.List(frmRegFiles.lstRegFiles.ListIndex))
Case "DX8VB.DLL":
    frmRegFiles.lblRegDesc.Caption = "The Microsoft DirectX8 dynamic library. Used to display graphics, play music, and run sound."
Exit Sub
Case "MSCOMM32.OCX":
    frmRegFiles.lblRegDesc.Caption = "A Microsoft control containing common controls used in Visual Basic 6.0 applications."
Exit Sub
Case "MSSTDFMT.DLL":
    frmRegFiles.lblRegDesc.Caption = "A Microsoft dynamic library containing necessary Visual Studio information."
Exit Sub
Case "MSWINSCK.OCX":
    frmRegFiles.lblRegDesc.Caption = "The Microsoft Winsock control. Used to transmit data between client and server programs."
Exit Sub
Case "RICHTX32.OCX":
    frmRegFiles.lblRegDesc.Caption = "The Microsoft Rich Text Box control. Enables Rich Text Boxes."
Exit Sub
Case "TABCTL32.OCX":
    frmRegFiles.lblRegDesc.Caption = "A Microsoft control."
Exit Sub
Case Else
    frmRegFiles.lblRegDesc.Caption = "No file description found!"
Exit Sub
End Select
End Sub
