VERSION 5.00
Begin VB.Form frmDroppingItem 
   Caption         =   "Dropping Items"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Item dropped into players inventory"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dropping Item on the Ground"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmDroppingItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmDroppingItem.Visible = False
End Sub

Private Sub Form_Load()
If GetVar(App.Path & "\Data.ini", "ADDED", "DroppingItem") = 0 Then
Option1.value = True
Option2.value = False
End If
If GetVar(App.Path & "\Data.ini", "ADDED", "DroppingItem") = 1 Then
Option1.value = False
Option2.value = True
End If
End Sub

Private Sub Option1_Click()
DroppingItem = 0
Call PutVar(App.Path & "\Data.ini", "ADDED", "DroppingItem", 0)
End Sub

Private Sub Option2_Click()
DroppingItem = 1
Call PutVar(App.Path & "\Data.ini", "ADDED", "DroppingItem", 1)
End Sub
