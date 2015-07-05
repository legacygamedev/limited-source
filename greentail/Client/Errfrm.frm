VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Errfrm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error"
   ClientHeight    =   5655
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "Errfrm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Would you like to send the error report to developer using E-mail?"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   6240
      TabIndex        =   14
      Top             =   4800
      Width           =   4815
      Begin VB.CommandButton nodonotsend 
         Caption         =   "Don't Send. "
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SendReport 
         Caption         =   "Send Report >>"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please, choose what you will like to do with program execution ?"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   6015
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit Application"
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         ToolTipText     =   "Abort current operation and terminate the application."
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abort Operation"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         ToolTipText     =   "Return to application and Abort Current Operation. Recommended"
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Continue Operation"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Return to application and Countinue What you were doing, Caution: May Produce Unpredicted Results."
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Debug Info:"
      Height          =   4455
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   4815
      Begin RichTextLib.RichTextBox RTDebug 
         Height          =   1455
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Debug Info For Developer"
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2566
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"Errfrm.frx":06C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Tcomment 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         ToolTipText     =   "Your Comments Here."
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your comments or questions below:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   4650
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "If you select to send error information to the developer, this information will be included with your E-mail. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   720
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Error Description:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6015
      Begin VB.Label lver 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "App. Exe. Info  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label edll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DLL Error No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Error No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label eno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Error Source :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label esrc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Error Description :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label edesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   5535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Error Has Ocurred"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   3780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Errfrm.frx":0745
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   6210
   End
End
Attribute VB_Name = "Errfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  VBErrorTrapDemo
'  How to Use:-
'  (1)Add Errfrm.frm, ModError.Bas, ModMSMail.bas _
'  Sysinfo.bas and ErrBitmap.res to your project.
'  (2)Add a refrence to Microsoft Scripting Runtime _
'  from Project --> references
'  (3)Add RichtextBox and winsock controls to your toolbox
'   from Project --> Components
'=========================================================================================
'  Coded By: Deepesh Agarwal
'  Published Date: 29/09/2003
'  WebSite: http://www.deepeshagarwal.tk
'  E-mail: agarwal_deepesh@indiatimes.com
'  Visit my site for Free-Software's like:
'  1). The-AdPolice - Blocks 17000+ adservers to save bandwidth
'  2). Dr. System -  Schedule Computer Maintainence - A must for every computer user
'  3). Service Controller XP (A Must For XP User) - Start,Stop,Pause and change startup type of 2000/XP services with recommended settings for different system config.
'   And Many More........
'=========================================================================================
Private Sub Form_Load()
    Dim ctr As Integer
    ctr = 101
End Sub

Private Sub nodonotsend_Click()
    'perform according to the selected radio option

    If Option1.Value = True Then 'Countinue

    End If
    If Option2.Value = True Then 'Abort

    End If
    If Option3.Value = True Then 'exit
        End
    End If
    Unload Me
End Sub

Private Sub SendReport_Click()
    Dim lRet As Long
    Dim Subject As String, sendto As String, msgBody As String
    Dim logname1 As String, logname2 As String, ForAttach As String
    Subject = "Error Debug Report for " & App.EXEName & " v " & App.Major & "." & App.Minor & "." & App.Revision
    sendto = "alex@scriptwisdom.com"
    logname1 = App.Path & "\ErrorLog.log"
    logname2 = App.Path & "\DebugLog.rtf"
    ForAttach = logname1 & ";" & logname2
    msgBody = Tcomment.Text
    msgBody = msgBody & vbCrLf & "Please Find Attached the Errlog.log and DebugLog.rtf Files"
    MsgBox "You will now be prompted by Outlook (Express) for confirmation. Please click yes to Queue this E-mail.", vbInformation, "Info"
    Call SendMail(Subject, sendto, sendto, ForAttach, msgBody)

    'perform according to the selected radio option

    If Option1.Value = True Then 'Countinue

    End If
    If Option2.Value = True Then 'Abort

    End If
    If Option3.Value = True Then 'exit
        End
    End If
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Static rid As Integer
    If rid = 0 Or rid > 103 Then
        rid = 101
    Else
        rid = rid + 1
    End If
End Sub
