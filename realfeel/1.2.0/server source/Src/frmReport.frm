VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtReport 
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
On Error GoTo ErrorHandler:
txtReport.Text = ""
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError Err.Source, "cmdClear_Click", Err.Number, Err.Description
End Sub
