VERSION 5.00
Begin VB.Form frmDrop 
   BorderStyle     =   0  'None
   Caption         =   "Drop Amount"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMinus1000 
      Caption         =   "- 1000"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPlus1000 
      Caption         =   "+ 1000"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus100 
      Caption         =   "- 100"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPlus100 
      Caption         =   "+ 100"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus10 
      Caption         =   "- 10"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPlus10 
      Caption         =   "+ 10"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus1 
      Caption         =   "- 1"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPlus1 
      Caption         =   "+ 1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblAmmount 
      Caption         =   "1"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Private Ammount As Long

Private Sub Form_Load()
Dim InvNum As Long

    Ammount = 1
    InvNum = frmCClient.lstPlayerInventory.ListIndex + 1
    
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 Then
        frmDrop.lblName = Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
    Else
        Unload Me
        frmCClient.picScreen.SetFocus
        Exit Sub
    End If
    Call ProcessAmmount
End Sub

Private Sub cmdOk_Click()
Dim InvNum As Long

    InvNum = frmCClient.lstPlayerInventory.ListIndex + 1
    
    Call SendDropItem(InvNum, Ammount)
    Unload Me
    frmCClient.picScreen.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    frmCClient.picScreen.SetFocus
End Sub

Private Sub cmdPlus1_Click()
    Ammount = Ammount + 1
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1_Click()
    Ammount = Ammount - 1
    Call ProcessAmmount
End Sub

Private Sub cmdPlus10_Click()
    Ammount = Ammount + 10
    Call ProcessAmmount
End Sub

Private Sub cmdMinus10_Click()
    Ammount = Ammount - 10
    Call ProcessAmmount
End Sub

Private Sub cmdPlus100_Click()
    Ammount = Ammount + 100
    Call ProcessAmmount
End Sub

Private Sub cmdMinus100_Click()
    Ammount = Ammount - 100
    Call ProcessAmmount
End Sub

Private Sub cmdPlus1000_Click()
    Ammount = Ammount + 1000
    Call ProcessAmmount
End Sub

Private Sub cmdMinus1000_Click()
    Ammount = Ammount - 1000
    Call ProcessAmmount
End Sub

Private Sub ProcessAmmount()
Dim InvNum As Long

    InvNum = frmCClient.lstPlayerInventory.ListIndex + 1
        
    ' Check if more then max and set back to max if so
    If Ammount > GetPlayerInvItemValue(MyIndex, InvNum) Then
        Ammount = GetPlayerInvItemValue(MyIndex, InvNum)
    End If
    
    ' Make sure its not 0
    If Ammount <= 0 Then
        Ammount = 1
    End If

    frmDrop.lblAmmount.Caption = Ammount & "/" & GetPlayerInvItemValue(MyIndex, InvNum)
End Sub

