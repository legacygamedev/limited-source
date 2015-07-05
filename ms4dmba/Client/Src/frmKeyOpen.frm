VERSION 5.00
Begin VB.Form frmKeyOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Open"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.HScrollBar scrlY 
      Height          =   255
      Left            =   480
      Max             =   11
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.HScrollBar scrlX 
      Height          =   255
      Left            =   480
      Max             =   15
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmKeyOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub cmdOk_Click()
    KeyOpenEditorX = scrlX.Value
    KeyOpenEditorY = scrlY.Value
    Unload Me
End Sub

Private Sub scrlX_Change()
    lblX.Caption = CStr(scrlX.Value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = CStr(scrlY.Value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
