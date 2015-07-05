VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label lblEvents42 
      Height          =   495
      Left            =   8280
      TabIndex        =   93
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents41 
      Height          =   495
      Left            =   6960
      TabIndex        =   92
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents40 
      Height          =   495
      Left            =   5640
      TabIndex        =   91
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents39 
      Height          =   495
      Left            =   4320
      TabIndex        =   90
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents38 
      Height          =   495
      Left            =   3000
      TabIndex        =   89
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents37 
      Height          =   495
      Left            =   1680
      TabIndex        =   88
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents36 
      Height          =   495
      Left            =   360
      TabIndex        =   87
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblEvents35 
      Height          =   495
      Left            =   8280
      TabIndex        =   86
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents34 
      Height          =   495
      Left            =   6960
      TabIndex        =   85
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents33 
      Height          =   495
      Left            =   5640
      TabIndex        =   84
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents32 
      Height          =   495
      Left            =   4320
      TabIndex        =   83
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents31 
      Height          =   495
      Left            =   3000
      TabIndex        =   82
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents30 
      Height          =   495
      Left            =   1680
      TabIndex        =   81
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents29 
      Height          =   495
      Left            =   360
      TabIndex        =   80
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblEvents28 
      Height          =   495
      Left            =   8280
      TabIndex        =   79
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents27 
      Height          =   495
      Left            =   6960
      TabIndex        =   78
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents26 
      Height          =   495
      Left            =   5640
      TabIndex        =   77
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents25 
      Height          =   495
      Left            =   4320
      TabIndex        =   76
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents24 
      Height          =   495
      Left            =   3000
      TabIndex        =   75
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents23 
      Height          =   495
      Left            =   1680
      TabIndex        =   74
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents22 
      Height          =   495
      Left            =   360
      TabIndex        =   73
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEvents21 
      Height          =   495
      Left            =   8280
      TabIndex        =   72
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents20 
      Height          =   495
      Left            =   6960
      TabIndex        =   71
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents19 
      Height          =   495
      Left            =   5640
      TabIndex        =   70
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents18 
      Height          =   495
      Left            =   4320
      TabIndex        =   69
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents17 
      Height          =   495
      Left            =   3000
      TabIndex        =   68
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents16 
      Height          =   495
      Left            =   1680
      TabIndex        =   67
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents15 
      Height          =   495
      Left            =   360
      TabIndex        =   66
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEvents14 
      Height          =   495
      Left            =   8280
      TabIndex        =   65
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents13 
      Height          =   495
      Left            =   6960
      TabIndex        =   64
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents12 
      Height          =   495
      Left            =   5640
      TabIndex        =   63
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents11 
      Height          =   495
      Left            =   4320
      TabIndex        =   62
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents10 
      Height          =   495
      Left            =   3000
      TabIndex        =   61
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents9 
      Height          =   495
      Left            =   1680
      TabIndex        =   60
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents8 
      Height          =   495
      Left            =   360
      TabIndex        =   59
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEvents7 
      Height          =   495
      Left            =   8280
      TabIndex        =   58
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEvents6 
      Height          =   495
      Left            =   6960
      TabIndex        =   57
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEvents5 
      Height          =   495
      Left            =   5640
      TabIndex        =   56
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEvents4 
      Height          =   495
      Left            =   4320
      TabIndex        =   55
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEvents3 
      Height          =   495
      Left            =   3000
      TabIndex        =   54
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lbl3 
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl2 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl1 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl42 
      Height          =   255
      Left            =   8280
      TabIndex        =   53
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl41 
      Height          =   255
      Left            =   6960
      TabIndex        =   52
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl40 
      Height          =   255
      Left            =   5640
      TabIndex        =   51
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl39 
      Height          =   255
      Left            =   4320
      TabIndex        =   50
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl38 
      Height          =   255
      Left            =   3000
      TabIndex        =   49
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl37 
      Height          =   255
      Left            =   1680
      TabIndex        =   48
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl36 
      Height          =   255
      Left            =   360
      TabIndex        =   47
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lbl35 
      Height          =   255
      Left            =   8280
      TabIndex        =   46
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl34 
      Height          =   255
      Left            =   6960
      TabIndex        =   45
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl33 
      Height          =   255
      Left            =   5640
      TabIndex        =   44
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl32 
      Height          =   255
      Left            =   4320
      TabIndex        =   43
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl31 
      Height          =   255
      Left            =   3000
      TabIndex        =   42
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lblSaturday 
      Alignment       =   2  'Center
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   41
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblFriday 
      Alignment       =   2  'Center
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   40
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblThursday 
      Alignment       =   2  'Center
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   39
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblWednesday 
      Alignment       =   2  'Center
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   38
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblTuesday 
      Alignment       =   2  'Center
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   37
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblMonday 
      Alignment       =   2  'Center
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   36
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblSunday 
      Alignment       =   2  'Center
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Line Line15 
      X1              =   9480
      X2              =   9480
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line14 
      X1              =   8160
      X2              =   8160
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Label lbl30 
      Height          =   255
      Left            =   1680
      TabIndex        =   34
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl29 
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lbl28 
      Height          =   255
      Left            =   8280
      TabIndex        =   32
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl27 
      Height          =   255
      Left            =   6960
      TabIndex        =   31
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl26 
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl25 
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl24 
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl23 
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl22 
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lbl21 
      Height          =   255
      Left            =   8280
      TabIndex        =   25
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl20 
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl19 
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl18 
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl17 
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl16 
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl15 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbl14 
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl13 
      Height          =   255
      Left            =   6960
      TabIndex        =   17
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl12 
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl11 
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl10 
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl9 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl8 
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lbl7 
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl6 
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl5 
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl4 
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblEvents2 
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Line Line13 
      X1              =   240
      X2              =   9480
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line12 
      X1              =   240
      X2              =   9480
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line11 
      X1              =   240
      X2              =   9480
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line10 
      X1              =   240
      X2              =   9480
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line9 
      X1              =   240
      X2              =   9480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   6840
      X2              =   6840
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line7 
      X1              =   5520
      X2              =   5520
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line6 
      X1              =   4200
      X2              =   4200
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line5 
      X1              =   2880
      X2              =   2880
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   9480
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   1560
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   2400
      Y2              =   8160
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Caption         =   "MonthName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9495
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label lblEvents1 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
year = 2005
month = 1
weekday = 2
day = 10

Me.lblYear = year

    Select Case month
        Case 1
            Me.lblMonth.Caption = "January"
        Case 2
            Me.lblMonth.Caption = "February"
        Case 3
            Me.lblMonth.Caption = "March"
        Case 4
            Me.lblMonth.Caption = "April"
        Case 5
            Me.lblMonth.Caption = "May"
        Case 6
            Me.lblMonth.Caption = "June"
        Case 7
            Me.lblMonth.Caption = "Juli"
        Case 8
            Me.lblMonth.Caption = "August"
        Case 9
            Me.lblMonth.Caption = "September"
        Case 10
            Me.lblMonth.Caption = "October"
        Case 11
            Me.lblMonth.Caption = "November"
        Case 12
            Me.lblMonth.Caption = "December"
    End Select
    
    Select Case weekday
        Case 1
            If day = 1 Then lblEvents2.ForeColor = vbBlue
            If day = 2 Then lblEvents3.ForeColor = vbBlue
            If day = 3 Then lblEvents4.ForeColor = vbBlue
            If day = 4 Then lblEvents5.ForeColor = vbBlue
            If day = 5 Then lblEvents6.ForeColor = vbBlue
            If day = 6 Then lblEvents7.ForeColor = vbBlue
            If day = 7 Then lblEvents8.ForeColor = vbBlue
            If day = 8 Then lblEvents9.ForeColor = vbBlue
            If day = 9 Then lblEvents10.ForeColor = vbBlue
            If day = 10 Then lblEvents11.ForeColor = vbBlue
            If day = 11 Then lblEvents12.ForeColor = vbBlue
            If day = 12 Then lblEvents13.ForeColor = vbBlue
            If day = 13 Then lblEvents14.ForeColor = vbBlue
            If day = 14 Then lblEvents15.ForeColor = vbBlue
            If day = 15 Then lblEvents16.ForeColor = vbBlue
            If day = 16 Then lblEvents17.ForeColor = vbBlue
            If day = 17 Then lblEvents18.ForeColor = vbBlue
            If day = 18 Then lblEvents19.ForeColor = vbBlue
            If day = 19 Then lblEvents20.ForeColor = vbBlue
            If day = 20 Then lblEvents21.ForeColor = vbBlue
            If day = 21 Then lblEvents22.ForeColor = vbBlue
            If day = 22 Then lblEvents23.ForeColor = vbBlue
            If day = 23 Then lblEvents24.ForeColor = vbBlue
            If day = 24 Then lblEvents25.ForeColor = vbBlue
            If day = 25 Then lblEvents26.ForeColor = vbBlue
            If day = 26 Then lblEvents27.ForeColor = vbBlue
            If day = 27 Then lblEvents28.ForeColor = vbBlue
            If day = 28 Then lblEvents29.ForeColor = vbBlue
            If day = 29 Then lblEvents30.ForeColor = vbBlue
            If day = 30 Then lblEvents31.ForeColor = vbBlue
            If day = 31 Then lblEvents32.ForeColor = vbBlue

            lblEvents2.Caption = "Day 1"
            lblEvents3.Caption = "Day 2"
            lblEvents4.Caption = "Day 3"
            lblEvents5.Caption = "Day 4"
            lblEvents6.Caption = "Day 5"
            lblEvents7.Caption = "Day 6"
            lblEvents8.Caption = "Day 7"
            lblEvents9.Caption = "Day 8"
            lblEvents10.Caption = "Day 9"
            lblEvents11.Caption = "Day 10"
            lblEvents12.Caption = "Day 11"
            lblEvents13.Caption = "Day 12"
            lblEvents14.Caption = "Day 13"
            lblEvents15.Caption = "Day 14"
            lblEvents16.Caption = "Day 15"
            lblEvents17.Caption = "Day 16"
            lblEvents18.Caption = "Day 17"
            lblEvents19.Caption = "Day 18"
            lblEvents20.Caption = "Day 19"
            lblEvents21.Caption = "Day 20"
            lblEvents22.Caption = "Day 21"
            lblEvents23.Caption = "Day 22"
            lblEvents24.Caption = "Day 23"
            lblEvents25.Caption = "Day 24"
            lblEvents26.Caption = "Day 25"
            lblEvents27.Caption = "Day 26"
            lblEvents28.Caption = "Day 27"
            lblEvents29.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents30.Caption = "Day 29"
            lblEvents31.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents32.Caption = "Day 31"
        Case 2
            If day = 1 Then lblEvents3.ForeColor = vbBlue
            If day = 2 Then lblEvents4.ForeColor = vbBlue
            If day = 3 Then lblEvents5.ForeColor = vbBlue
            If day = 4 Then lblEvents6.ForeColor = vbBlue
            If day = 5 Then lblEvents7.ForeColor = vbBlue
            If day = 6 Then lblEvents8.ForeColor = vbBlue
            If day = 7 Then lblEvents9.ForeColor = vbBlue
            If day = 8 Then lblEvents10.ForeColor = vbBlue
            If day = 9 Then lblEvents11.ForeColor = vbBlue
            If day = 10 Then lblEvents12.ForeColor = vbBlue
            If day = 11 Then lblEvents13.ForeColor = vbBlue
            If day = 12 Then lblEvents14.ForeColor = vbBlue
            If day = 13 Then lblEvents15.ForeColor = vbBlue
            If day = 14 Then lblEvents16.ForeColor = vbBlue
            If day = 15 Then lblEvents17.ForeColor = vbBlue
            If day = 16 Then lblEvents18.ForeColor = vbBlue
            If day = 17 Then lblEvents19.ForeColor = vbBlue
            If day = 18 Then lblEvents20.ForeColor = vbBlue
            If day = 19 Then lblEvents21.ForeColor = vbBlue
            If day = 20 Then lblEvents22.ForeColor = vbBlue
            If day = 21 Then lblEvents23.ForeColor = vbBlue
            If day = 22 Then lblEvents24.ForeColor = vbBlue
            If day = 23 Then lblEvents25.ForeColor = vbBlue
            If day = 24 Then lblEvents26.ForeColor = vbBlue
            If day = 25 Then lblEvents27.ForeColor = vbBlue
            If day = 26 Then lblEvents28.ForeColor = vbBlue
            If day = 27 Then lblEvents29.ForeColor = vbBlue
            If day = 28 Then lblEvents30.ForeColor = vbBlue
            If day = 29 Then lblEvents31.ForeColor = vbBlue
            If day = 30 Then lblEvents32.ForeColor = vbBlue
            If day = 31 Then lblEvents33.ForeColor = vbBlue
            
            lblEvents3.Caption = "Day 1"
            lblEvents4.Caption = "Day 2"
            lblEvents5.Caption = "Day 3"
            lblEvents6.Caption = "Day 4"
            lblEvents7.Caption = "Day 5"
            lblEvents8.Caption = "Day 6"
            lblEvents9.Caption = "Day 7"
            lblEvents10.Caption = "Day 8"
            lblEvents11.Caption = "Day 9"
            lblEvents12.Caption = "Day 10"
            lblEvents13.Caption = "Day 11"
            lblEvents14.Caption = "Day 12"
            lblEvents15.Caption = "Day 13"
            lblEvents16.Caption = "Day 14"
            lblEvents17.Caption = "Day 15"
            lblEvents18.Caption = "Day 16"
            lblEvents19.Caption = "Day 17"
            lblEvents20.Caption = "Day 18"
            lblEvents21.Caption = "Day 19"
            lblEvents22.Caption = "Day 20"
            lblEvents23.Caption = "Day 21"
            lblEvents24.Caption = "Day 22"
            lblEvents25.Caption = "Day 23"
            lblEvents26.Caption = "Day 24"
            lblEvents27.Caption = "Day 25"
            lblEvents28.Caption = "Day 26"
            lblEvents29.Caption = "Day 27"
            lblEvents30.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents31.Caption = "Day 29"
            lblEvents32.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents33.Caption = "Day 31"
        Case 3
            If day = 1 Then lblEvents4.ForeColor = vbBlue
            If day = 2 Then lblEvents5.ForeColor = vbBlue
            If day = 3 Then lblEvents6.ForeColor = vbBlue
            If day = 4 Then lblEvents7.ForeColor = vbBlue
            If day = 5 Then lblEvents8.ForeColor = vbBlue
            If day = 6 Then lblEvents9.ForeColor = vbBlue
            If day = 7 Then lblEvents10.ForeColor = vbBlue
            If day = 8 Then lblEvents11.ForeColor = vbBlue
            If day = 9 Then lblEvents12.ForeColor = vbBlue
            If day = 10 Then lblEvents13.ForeColor = vbBlue
            If day = 11 Then lblEvents14.ForeColor = vbBlue
            If day = 12 Then lblEvents15.ForeColor = vbBlue
            If day = 13 Then lblEvents16.ForeColor = vbBlue
            If day = 14 Then lblEvents17.ForeColor = vbBlue
            If day = 15 Then lblEvents18.ForeColor = vbBlue
            If day = 16 Then lblEvents19.ForeColor = vbBlue
            If day = 17 Then lblEvents20.ForeColor = vbBlue
            If day = 18 Then lblEvents21.ForeColor = vbBlue
            If day = 19 Then lblEvents22.ForeColor = vbBlue
            If day = 20 Then lblEvents23.ForeColor = vbBlue
            If day = 21 Then lblEvents24.ForeColor = vbBlue
            If day = 22 Then lblEvents25.ForeColor = vbBlue
            If day = 23 Then lblEvents26.ForeColor = vbBlue
            If day = 24 Then lblEvents27.ForeColor = vbBlue
            If day = 25 Then lblEvents28.ForeColor = vbBlue
            If day = 26 Then lblEvents29.ForeColor = vbBlue
            If day = 27 Then lblEvents30.ForeColor = vbBlue
            If day = 28 Then lblEvents31.ForeColor = vbBlue
            If day = 29 Then lblEvents32.ForeColor = vbBlue
            If day = 30 Then lblEvents33.ForeColor = vbBlue
            If day = 31 Then lblEvents34.ForeColor = vbBlue
            
            lblEvents4.Caption = "Day 1"
            lblEvents5.Caption = "Day 2"
            lblEvents6.Caption = "Day 3"
            lblEvents7.Caption = "Day 4"
            lblEvents8.Caption = "Day 5"
            lblEvents9.Caption = "Day 6"
            lblEvents10.Caption = "Day 7"
            lblEvents11.Caption = "Day 8"
            lblEvents12.Caption = "Day 9"
            lblEvents13.Caption = "Day 10"
            lblEvents14.Caption = "Day 11"
            lblEvents15.Caption = "Day 12"
            lblEvents16.Caption = "Day 13"
            lblEvents17.Caption = "Day 14"
            lblEvents18.Caption = "Day 15"
            lblEvents19.Caption = "Day 16"
            lblEvents20.Caption = "Day 17"
            lblEvents21.Caption = "Day 18"
            lblEvents22.Caption = "Day 19"
            lblEvents23.Caption = "Day 20"
            lblEvents24.Caption = "Day 21"
            lblEvents25.Caption = "Day 22"
            lblEvents26.Caption = "Day 23"
            lblEvents27.Caption = "Day 24"
            lblEvents28.Caption = "Day 25"
            lblEvents29.Caption = "Day 26"
            lblEvents30.Caption = "Day 27"
            lblEvents31.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents32.Caption = "Day 29"
            lblEvents33.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents34.Caption = "Day 31"
        Case 4
            If day = 1 Then lblEvents5.ForeColor = vbBlue
            If day = 2 Then lblEvents6.ForeColor = vbBlue
            If day = 3 Then lblEvents7.ForeColor = vbBlue
            If day = 4 Then lblEvents8.ForeColor = vbBlue
            If day = 5 Then lblEvents9.ForeColor = vbBlue
            If day = 6 Then lblEvents10.ForeColor = vbBlue
            If day = 7 Then lblEvents11.ForeColor = vbBlue
            If day = 8 Then lblEvents12.ForeColor = vbBlue
            If day = 9 Then lblEvents13.ForeColor = vbBlue
            If day = 10 Then lblEvents14.ForeColor = vbBlue
            If day = 11 Then lblEvents15.ForeColor = vbBlue
            If day = 12 Then lblEvents16.ForeColor = vbBlue
            If day = 13 Then lblEvents17.ForeColor = vbBlue
            If day = 14 Then lblEvents18.ForeColor = vbBlue
            If day = 15 Then lblEvents19.ForeColor = vbBlue
            If day = 16 Then lblEvents20.ForeColor = vbBlue
            If day = 17 Then lblEvents21.ForeColor = vbBlue
            If day = 18 Then lblEvents22.ForeColor = vbBlue
            If day = 19 Then lblEvents23.ForeColor = vbBlue
            If day = 20 Then lblEvents24.ForeColor = vbBlue
            If day = 21 Then lblEvents25.ForeColor = vbBlue
            If day = 22 Then lblEvents26.ForeColor = vbBlue
            If day = 23 Then lblEvents27.ForeColor = vbBlue
            If day = 24 Then lblEvents28.ForeColor = vbBlue
            If day = 25 Then lblEvents29.ForeColor = vbBlue
            If day = 26 Then lblEvents30.ForeColor = vbBlue
            If day = 27 Then lblEvents31.ForeColor = vbBlue
            If day = 28 Then lblEvents32.ForeColor = vbBlue
            If day = 29 Then lblEvents33.ForeColor = vbBlue
            If day = 30 Then lblEvents34.ForeColor = vbBlue
            If day = 31 Then lblEvents35.ForeColor = vbBlue
            
            lblEvents5.Caption = "Day 1"
            lblEvents6.Caption = "Day 2"
            lblEvents7.Caption = "Day 3"
            lblEvents8.Caption = "Day 4"
            lblEvents9.Caption = "Day 5"
            lblEvents10.Caption = "Day 6"
            lblEvents11.Caption = "Day 7"
            lblEvents12.Caption = "Day 8"
            lblEvents13.Caption = "Day 9"
            lblEvents14.Caption = "Day 10"
            lblEvents15.Caption = "Day 11"
            lblEvents16.Caption = "Day 12"
            lblEvents17.Caption = "Day 13"
            lblEvents18.Caption = "Day 14"
            lblEvents19.Caption = "Day 15"
            lblEvents20.Caption = "Day 16"
            lblEvents21.Caption = "Day 17"
            lblEvents22.Caption = "Day 18"
            lblEvents23.Caption = "Day 19"
            lblEvents24.Caption = "Day 20"
            lblEvents25.Caption = "Day 21"
            lblEvents26.Caption = "Day 22"
            lblEvents27.Caption = "Day 23"
            lblEvents28.Caption = "Day 24"
            lblEvents29.Caption = "Day 25"
            lblEvents30.Caption = "Day 26"
            lblEvents31.Caption = "Day 27"
            lblEvents32.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents33.Caption = "Day 29"
            lblEvents34.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents35.Caption = "Day 31"
        Case 5
            If day = 1 Then lblEvents6.ForeColor = vbBlue
            If day = 2 Then lblEvents7.ForeColor = vbBlue
            If day = 3 Then lblEvents8.ForeColor = vbBlue
            If day = 4 Then lblEvents9.ForeColor = vbBlue
            If day = 5 Then lblEvents10.ForeColor = vbBlue
            If day = 6 Then lblEvents11.ForeColor = vbBlue
            If day = 7 Then lblEvents12.ForeColor = vbBlue
            If day = 8 Then lblEvents13.ForeColor = vbBlue
            If day = 9 Then lblEvents14.ForeColor = vbBlue
            If day = 10 Then lblEvents15.ForeColor = vbBlue
            If day = 11 Then lblEvents16.ForeColor = vbBlue
            If day = 12 Then lblEvents17.ForeColor = vbBlue
            If day = 13 Then lblEvents18.ForeColor = vbBlue
            If day = 14 Then lblEvents19.ForeColor = vbBlue
            If day = 15 Then lblEvents20.ForeColor = vbBlue
            If day = 16 Then lblEvents21.ForeColor = vbBlue
            If day = 17 Then lblEvents22.ForeColor = vbBlue
            If day = 18 Then lblEvents23.ForeColor = vbBlue
            If day = 19 Then lblEvents24.ForeColor = vbBlue
            If day = 20 Then lblEvents25.ForeColor = vbBlue
            If day = 21 Then lblEvents26.ForeColor = vbBlue
            If day = 22 Then lblEvents27.ForeColor = vbBlue
            If day = 23 Then lblEvents28.ForeColor = vbBlue
            If day = 24 Then lblEvents29.ForeColor = vbBlue
            If day = 25 Then lblEvents30.ForeColor = vbBlue
            If day = 26 Then lblEvents31.ForeColor = vbBlue
            If day = 27 Then lblEvents32.ForeColor = vbBlue
            If day = 28 Then lblEvents33.ForeColor = vbBlue
            If day = 29 Then lblEvents34.ForeColor = vbBlue
            If day = 30 Then lblEvents35.ForeColor = vbBlue
            If day = 31 Then lblEvents36.ForeColor = vbBlue
        
            lblEvents6.Caption = "Day 1"
            lblEvents7.Caption = "Day 2"
            lblEvents8.Caption = "Day 3"
            lblEvents9.Caption = "Day 4"
            lblEvents10.Caption = "Day 5"
            lblEvents11.Caption = "Day 6"
            lblEvents12.Caption = "Day 7"
            lblEvents13.Caption = "Day 8"
            lblEvents14.Caption = "Day 9"
            lblEvents15.Caption = "Day 10"
            lblEvents16.Caption = "Day 11"
            lblEvents17.Caption = "Day 12"
            lblEvents18.Caption = "Day 13"
            lblEvents19.Caption = "Day 14"
            lblEvents20.Caption = "Day 15"
            lblEvents21.Caption = "Day 16"
            lblEvents22.Caption = "Day 17"
            lblEvents23.Caption = "Day 18"
            lblEvents24.Caption = "Day 19"
            lblEvents25.Caption = "Day 20"
            lblEvents26.Caption = "Day 21"
            lblEvents27.Caption = "Day 22"
            lblEvents28.Caption = "Day 23"
            lblEvents29.Caption = "Day 24"
            lblEvents30.Caption = "Day 25"
            lblEvents31.Caption = "Day 26"
            lblEvents32.Caption = "Day 27"
            lblEvents33.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents34.Caption = "Day 29"
            lblEvents35.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents36.Caption = "Day 31"
        Case 6
            If day = 1 Then lblEvents7.ForeColor = vbBlue
            If day = 2 Then lblEvents8.ForeColor = vbBlue
            If day = 3 Then lblEvents9.ForeColor = vbBlue
            If day = 4 Then lblEvents10.ForeColor = vbBlue
            If day = 5 Then lblEvents11.ForeColor = vbBlue
            If day = 6 Then lblEvents12.ForeColor = vbBlue
            If day = 7 Then lblEvents13.ForeColor = vbBlue
            If day = 8 Then lblEvents14.ForeColor = vbBlue
            If day = 9 Then lblEvents15.ForeColor = vbBlue
            If day = 10 Then lblEvents16.ForeColor = vbBlue
            If day = 11 Then lblEvents17.ForeColor = vbBlue
            If day = 12 Then lblEvents18.ForeColor = vbBlue
            If day = 13 Then lblEvents19.ForeColor = vbBlue
            If day = 14 Then lblEvents20.ForeColor = vbBlue
            If day = 15 Then lblEvents21.ForeColor = vbBlue
            If day = 16 Then lblEvents22.ForeColor = vbBlue
            If day = 17 Then lblEvents23.ForeColor = vbBlue
            If day = 18 Then lblEvents24.ForeColor = vbBlue
            If day = 19 Then lblEvents25.ForeColor = vbBlue
            If day = 20 Then lblEvents26.ForeColor = vbBlue
            If day = 21 Then lblEvents27.ForeColor = vbBlue
            If day = 22 Then lblEvents28.ForeColor = vbBlue
            If day = 23 Then lblEvents29.ForeColor = vbBlue
            If day = 24 Then lblEvents30.ForeColor = vbBlue
            If day = 25 Then lblEvents31.ForeColor = vbBlue
            If day = 26 Then lblEvents32.ForeColor = vbBlue
            If day = 27 Then lblEvents33.ForeColor = vbBlue
            If day = 28 Then lblEvents34.ForeColor = vbBlue
            If day = 29 Then lblEvents35.ForeColor = vbBlue
            If day = 30 Then lblEvents36.ForeColor = vbBlue
            If day = 31 Then lblEvents37.ForeColor = vbBlue

            lblEvents7.Caption = "Day 1"
            lblEvents8.Caption = "Day 2"
            lblEvents9.Caption = "Day 3"
            lblEvents10.Caption = "Day 4"
            lblEvents11.Caption = "Day 5"
            lblEvents12.Caption = "Day 6"
            lblEvents13.Caption = "Day 7"
            lblEvents14.Caption = "Day 8"
            lblEvents15.Caption = "Day 9"
            lblEvents16.Caption = "Day 10"
            lblEvents17.Caption = "Day 11"
            lblEvents18.Caption = "Day 12"
            lblEvents19.Caption = "Day 13"
            lblEvents20.Caption = "Day 14"
            lblEvents21.Caption = "Day 15"
            lblEvents22.Caption = "Day 16"
            lblEvents23.Caption = "Day 17"
            lblEvents24.Caption = "Day 18"
            lblEvents25.Caption = "Day 19"
            lblEvents26.Caption = "Day 20"
            lblEvents27.Caption = "Day 21"
            lblEvents28.Caption = "Day 22"
            lblEvents29.Caption = "Day 23"
            lblEvents30.Caption = "Day 24"
            lblEvents31.Caption = "Day 25"
            lblEvents32.Caption = "Day 26"
            lblEvents33.Caption = "Day 27"
            lblEvents34.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents35.Caption = "Day 29"
            lblEvents36.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents37.Caption = "Day 31"
        Case 7
            If day = 1 Then lblEvents1.ForeColor = vbBlue
            If day = 2 Then lblEvents2.ForeColor = vbBlue
            If day = 3 Then lblEvents3.ForeColor = vbBlue
            If day = 4 Then lblEvents4.ForeColor = vbBlue
            If day = 5 Then lblEvents5.ForeColor = vbBlue
            If day = 6 Then lblEvents6.ForeColor = vbBlue
            If day = 7 Then lblEvents7.ForeColor = vbBlue
            If day = 8 Then lblEvents8.ForeColor = vbBlue
            If day = 9 Then lblEvents9.ForeColor = vbBlue
            If day = 10 Then lblEvents10.ForeColor = vbBlue
            If day = 11 Then lblEvents11.ForeColor = vbBlue
            If day = 12 Then lblEvents12.ForeColor = vbBlue
            If day = 13 Then lblEvents13.ForeColor = vbBlue
            If day = 14 Then lblEvents14.ForeColor = vbBlue
            If day = 15 Then lblEvents15.ForeColor = vbBlue
            If day = 16 Then lblEvents16.ForeColor = vbBlue
            If day = 17 Then lblEvents17.ForeColor = vbBlue
            If day = 18 Then lblEvents18.ForeColor = vbBlue
            If day = 19 Then lblEvents19.ForeColor = vbBlue
            If day = 20 Then lblEvents20.ForeColor = vbBlue
            If day = 21 Then lblEvents21.ForeColor = vbBlue
            If day = 22 Then lblEvents22.ForeColor = vbBlue
            If day = 23 Then lblEvents23.ForeColor = vbBlue
            If day = 24 Then lblEvents24.ForeColor = vbBlue
            If day = 25 Then lblEvents25.ForeColor = vbBlue
            If day = 26 Then lblEvents26.ForeColor = vbBlue
            If day = 27 Then lblEvents27.ForeColor = vbBlue
            If day = 28 Then lblEvents28.ForeColor = vbBlue
            If day = 29 Then lblEvents29.ForeColor = vbBlue
            If day = 30 Then lblEvents30.ForeColor = vbBlue
            If day = 31 Then lblEvents31.ForeColor = vbBlue
        
            lblEvents1.Caption = "Day 1"
            lblEvents2.Caption = "Day 2"
            lblEvents3.Caption = "Day 3"
            lblEvents4.Caption = "Day 4"
            lblEvents5.Caption = "Day 5"
            lblEvents6.Caption = "Day 6"
            lblEvents7.Caption = "Day 7"
            lblEvents8.Caption = "Day 8"
            lblEvents9.Caption = "Day 9"
            lblEvents10.Caption = "Day 10"
            lblEvents11.Caption = "Day 11"
            lblEvents12.Caption = "Day 12"
            lblEvents13.Caption = "Day 13"
            lblEvents14.Caption = "Day 14"
            lblEvents15.Caption = "Day 15"
            lblEvents16.Caption = "Day 16"
            lblEvents17.Caption = "Day 17"
            lblEvents18.Caption = "Day 18"
            lblEvents19.Caption = "Day 19"
            lblEvents20.Caption = "Day 20"
            lblEvents21.Caption = "Day 21"
            lblEvents22.Caption = "Day 22"
            lblEvents23.Caption = "Day 23"
            lblEvents24.Caption = "Day 24"
            lblEvents25.Caption = "Day 25"
            lblEvents26.Caption = "Day 26"
            lblEvents27.Caption = "Day 27"
            lblEvents28.Caption = "Day 28"
            If month = 2 Then Exit Sub
            lblEvents29.Caption = "Day 29"
            lblEvents30.Caption = "Day 30"
            If month = 4 Or month = 6 Or month = 8 Or month = 10 Or month = 12 Then Exit Sub
            lblEvents31.Caption = "Day 31"
    End Select
End Sub
