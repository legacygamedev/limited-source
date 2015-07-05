VERSION 5.00
Begin VB.Form frmQuestEditor 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGoldGiven 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   855
      TabIndex        =   22
      Text            =   "0"
      Top             =   5070
      Width           =   1155
   End
   Begin VB.TextBox txtMinLevel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3660
      TabIndex        =   21
      Text            =   "0"
      Top             =   4755
      Width           =   1155
   End
   Begin VB.TextBox txtexp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   810
      TabIndex        =   19
      Text            =   "0"
      Top             =   4755
      Width           =   1155
   End
   Begin VB.TextBox txtItemVal2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4155
      TabIndex        =   17
      Text            =   "1"
      Top             =   4410
      Width           =   780
   End
   Begin VB.TextBox txtItemDesc2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4410
      Width           =   3135
   End
   Begin VB.TextBox txtItemNum2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3135
      TabIndex        =   13
      Text            =   "0"
      Top             =   4410
      Width           =   570
   End
   Begin VB.HScrollBar scrItemGiv 
      Height          =   225
      Left            =   0
      Max             =   1020
      TabIndex        =   12
      Top             =   4200
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1365
      TabIndex        =   10
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   5460
      Width           =   1365
   End
   Begin VB.OptionButton optFinishText 
      Caption         =   "Finish Quest Text"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   3060
      Width           =   1530
   End
   Begin VB.OptionButton optMidText 
      Caption         =   "After Item Found Text"
      Height          =   195
      Left            =   1500
      TabIndex        =   7
      Top             =   3060
      Width           =   1830
   End
   Begin VB.OptionButton optStartText 
      Caption         =   "Start Quest Text"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   3060
      Value           =   -1  'True
      Width           =   1530
   End
   Begin VB.TextBox txtItemDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3705
      Width           =   4275
   End
   Begin VB.TextBox txtItemNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4260
      TabIndex        =   4
      Text            =   "0"
      Top             =   3705
      Width           =   660
   End
   Begin VB.HScrollBar scrItemCol 
      Height          =   225
      Left            =   -15
      Max             =   1020
      TabIndex        =   3
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox txtQuestFinishText 
      Appearance      =   0  'Flat
      Height          =   3000
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   15
      Width           =   4905
   End
   Begin VB.TextBox txtQuestMiddleText 
      Appearance      =   0  'Flat
      Height          =   3000
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   15
      Width           =   4905
   End
   Begin VB.TextBox txtQuestStartText 
      Appearance      =   0  'Flat
      Height          =   3000
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   15
      Width           =   4905
   End
   Begin VB.Label Label6 
      Caption         =   "Gold given:"
      Height          =   315
      Left            =   15
      TabIndex        =   23
      Top             =   5085
      Width           =   810
   End
   Begin VB.Label Label5 
      Caption         =   "Required Min Level:"
      Height          =   285
      Left            =   2190
      TabIndex        =   20
      Top             =   4770
      Width           =   2715
   End
   Begin VB.Label Label4 
      Caption         =   "Exp given:"
      Height          =   315
      Left            =   15
      TabIndex        =   18
      Top             =   4770
      Width           =   810
   End
   Begin VB.Label Label3 
      Caption         =   "Value"
      Height          =   210
      Left            =   3720
      TabIndex        =   16
      Top             =   4455
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Item Given"
      Height          =   240
      Left            =   15
      TabIndex        =   15
      Top             =   3990
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Item player must collect"
      Height          =   210
      Left            =   45
      TabIndex        =   11
      Top             =   3285
      Width           =   3150
   End
End
Attribute VB_Name = "frmQuestEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call questEditorCancel
End Sub

Private Sub cmdSave_Click()
    Call QuestEditorOk
End Sub

Private Sub Form_Load()
    txtQuestStartText.Visible = True
    txtQuestMiddleText.Visible = False
    txtQuestFinishText.Visible = False
End Sub

Private Sub optFinishText_Click()
    txtQuestStartText.Visible = False
    txtQuestMiddleText.Visible = False
    txtQuestFinishText.Visible = True
End Sub

Private Sub optMidText_Click()
    txtQuestStartText.Visible = False
    txtQuestMiddleText.Visible = True
    txtQuestFinishText.Visible = False
End Sub

Private Sub optStartText_Click()
    txtQuestStartText.Visible = True
    txtQuestMiddleText.Visible = False
    txtQuestFinishText.Visible = False
End Sub

Private Sub scrItemCol_Change()
On Error Resume Next
    Me.txtItemDesc = Item(scrItemCol).name
    Me.txtItemNo = scrItemCol
End Sub

Private Sub scrItemGiv_Change()
On Error Resume Next
    Me.txtItemDesc2 = Item(scrItemGiv).name
    Me.txtItemNum2 = scrItemGiv
End Sub
