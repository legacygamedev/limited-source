VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmArchive 
   Caption         =   "Archiver"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   3975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   8055
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8055
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
    CD.ShowOpen
    CD.FileName
End Sub
