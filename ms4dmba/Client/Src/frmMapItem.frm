VERSION 5.00
Begin VB.Form frmMapItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Item"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
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
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   50
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.HScrollBar scrlValue 
      Height          =   255
      Left            =   840
      Min             =   1
      TabIndex        =   2
      Top             =   960
      Value           =   1
      Width           =   3255
   End
   Begin VB.HScrollBar scrlItem 
      Height          =   255
      Left            =   840
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   600
      Value           =   1
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMapItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
    MapItemEditorBltItem
    scrlItem.Max = MAX_ITEMS
    lblName.Caption = Trim$(Item(scrlItem.Value).Name)
    scrlItem_Change
End Sub

Private Sub cmdOk_Click()
    ItemEditorNum = scrlItem.Value
    ItemEditorValue = scrlValue.Value
    Unload Me
End Sub

Private Sub scrlItem_Change()
    MapItemEditorBltItem
    lblItem.Caption = CStr(scrlItem.Value)
    lblName.Caption = Trim$(Item(scrlItem.Value).Name)
    
    If Item(scrlItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlValue.Enabled = True
    Else
        scrlValue.Enabled = False
        lblValue = 1
    End If
    
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = CStr(scrlValue.Value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

