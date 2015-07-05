VERSION 5.00
Begin VB.Form frmDesign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu Design Tool"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   636
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   924
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDesignEditor 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   10800
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   135
         Index           =   0
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Frame fraDesignSelected 
         Caption         =   "Selected"
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton cmdDesignApply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   1320
            TabIndex        =   77
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdDesignPreview 
            Caption         =   "Preview"
            Height          =   375
            Left            =   240
            TabIndex        =   76
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtDesignWidth 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDesignHeight 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   11
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdDesignPrevious 
            Caption         =   "Previous"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdDesignNext 
            Caption         =   "Next"
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblDesignPicNum 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   1080
            TabIndex        =   8
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Width"
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Height"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "New Character"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   3000
         Width           =   1455
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Player Characters"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   26
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Credits"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   2520
         Width           =   855
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Delete Account"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "New Account"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Login"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton optDesignOption 
         Caption         =   "Main Menu"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   1560
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdDesignCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1500
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDesignSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtDesignDesigner 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtDesignName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Designed By"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Design Name"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox picDesignBackground 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6375
      Index           =   1
      Left            =   240
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   591
      TabIndex        =   15
      Top             =   240
      Width           =   8895
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   7
         Left            =   2400
         ScaleHeight     =   119
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   407
         TabIndex        =   41
         Top             =   4320
         Visible         =   0   'False
         Width           =   6135
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   7
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignNewChar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Name Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   3120
            TabIndex        =   60
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Dexterity Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   8
            Left            =   1440
            TabIndex        =   58
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Speed Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   7
            Left            =   1440
            TabIndex        =   57
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Magic Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   6
            Left            =   1440
            TabIndex        =   56
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Defence Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   5
            Left            =   1440
            TabIndex        =   55
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Strength Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   1440
            TabIndex        =   54
            Top             =   120
            Width           =   1500
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - SP Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   53
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - MP Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - HP Value"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Female Option"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   4560
            TabIndex        =   50
            Top             =   840
            Width           =   1350
         End
         Begin VB.Label lblDesignNewChar 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0 - Male Option"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   3120
            TabIndex        =   49
            Top             =   840
            Width           =   1350
         End
         Begin VB.Label lblDesignNewChar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Class Dropdown Menu"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   48
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblDesignNewChar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   3720
            TabIndex        =   47
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDesignNewChar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Add Character"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   960
            TabIndex        =   46
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDesignNewChar 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   6
         Left            =   4080
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   6
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignChars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Character List"
            ForeColor       =   &H80000008&
            Height          =   1095
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblDesignChars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Quit"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   45
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblDesignChars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Delete Character"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2160
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblDesignChars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Use Character"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   43
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblDesignChars 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "New Character"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   42
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblDesignChars 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   5
         Left            =   4080
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   5
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignCredits 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Credits text here......"
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   1
            Left            =   240
            TabIndex        =   39
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label lblDesignCredits 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   38
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblDesignCredits 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   4
         Left            =   4080
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   4
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignDelAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Password Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   71
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblDesignDelAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Name Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   70
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblDesignDelAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   36
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label lblDesignDelAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   35
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblDesignDelAcc 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   3
         Left            =   4080
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   3
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignNewAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Password Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   68
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblDesignNewAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Name Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   67
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblDesignNewAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   33
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label lblDesignNewAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   32
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblDesignNewAcc 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox picDesignBackground 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   2
         Left            =   4080
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Shape shpDesignSelect 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   975
            Index           =   2
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDesignLogin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Password Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   65
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblDesignLogin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Name Entry"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   64
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label lblDesignLogin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cancel"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   30
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label lblDesignLogin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   29
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblDesignLogin 
            BackStyle       =   0  'Transparent
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.Label lblDesignMenu 
         Height          =   135
         Index           =   0
         Left            =   1200
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Shape shpDesignSelect 
         BorderColor     =   &H00000080&
         BorderWidth     =   4
         Height          =   975
         Index           =   1
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDesignMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quit"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   480
         TabIndex        =   20
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label lblDesignMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Credits"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label lblDesignMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete Account"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblDesignMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New Account"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblDesignMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Login"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.PictureBox picDesignBackground 
      Height          =   1215
      Index           =   0
      Left            =   1320
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape shpDesignSelect 
      BorderColor     =   &H00000080&
      BorderWidth     =   4
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmDesign"
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

' *Backgrounds*              *Menu Components*             *Login Components*
' 1 = Menu                   1 = Login label               1 = Name textbox
' 2 = Login                  2 = New Account label         2 = Password textbox
' 3 = New Account            3 = Delete Account label      3 = OK button
' 4 = Delete Account         4 = Credits label             4 = Cancel button
' 5 = Credits                5 = Quit label
' 6 = Character List
' 7 = Add Character
'
' *New Account Components*   *Delete Account Components*   *Credits Components*
' 1 = Name textbox           1 = Name textbox              1 = Credits textbox
' 2 = Password textbox       2 = Password textbox          2 = Cancel button
' 3 = OK button              3 = OK button
' 4 = Cancel button          4 = Cancel button
'
' *Character List*              *Add Character Components*
' 1 = Character listbox         1 = HP Value label        8 = Dexterity Value label
' 2 = New Character label       2 = MP Value label        9 = Name Entry textbox
' 3 = Use Character label       3 = SP Value label        10 = Class dropdown
' 4 = Delete Character label    4 = Strength Value label  11 = Male Option button
' 5 = Quit label                5 = Defence Value label   12 = Female Option button
'                               6 = Magic Value label     13 = Add Character button
'                               7 = Speed Value label     14 = Cancel button

Private Type DesignSelectRec
    Background As Byte
    Control As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
End Type

Private Selection As DesignSelectRec

Private Selected As Boolean
Private Changed As Boolean

Private Sub Form_Load()
Dim i As Long

    Me.ScaleMode = 3
    Call ResetSelection
    For i = 0 To 7
        shpDesignSelect(i).Visible = False
    Next i
    cmdDesignNext.Enabled = False
    cmdDesignPrevious.Enabled = False
    cmdDesignPreview.Enabled = False
    cmdDesignApply.Enabled = False
    Changed = False
    frmCClient.picScreen.Refresh
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    If Selected = False Then
        Source.Move x - MouseXOffset, y - MouseYOffset
        Source.Visible = True
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call ResetChanges
        Call ResetSelection
    End If
End Sub

' *******************
' ** Design Editor **
' *******************

Private Sub picDesignEditor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call ResetChanges
        Call ResetSelection
    Else
        If optDesignOption(1).Value = True Then
            MouseXOffset = x
            MouseYOffset = y
            picDesignEditor.Drag vbBeginDrag
        End If
    End If
End Sub

Private Sub cmdDesignNext_Click()
    If Selection.Control = 0 Then
        If Selection.Background > 0 Then
            picDesignBackground(Selection.Background).Picture = LoadPicture(App.Path & DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption + 1) & GFX_EXT)
            lblDesignPicNum.Caption = (lblDesignPicNum.Caption + 1)
            Selection.Data5 = Val(lblDesignPicNum.Caption)
            If FileExist(DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption + 1) & GFX_EXT, False) Then
                cmdDesignNext.Enabled = True
            Else
                cmdDesignNext.Enabled = False
            End If
            If FileExist(DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption - 1) & GFX_EXT, False) Or (lblDesignPicNum.Caption - 1 = 0) Then
                cmdDesignPrevious.Enabled = True
            Else
                cmdDesignPrevious.Enabled = False
            End If
            Changed = True
            cmdDesignApply.Enabled = True
        End If
    End If
End Sub

Private Sub cmdDesignPrevious_Click()
    If Selection.Control = 0 Then
        If (lblDesignPicNum.Caption - 1) > 0 Then
            picDesignBackground(Selection.Background).Picture = LoadPicture(App.Path & DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption - 1) & GFX_EXT)
        Else
            picDesignBackground(Selection.Background).Picture = LoadPicture()
        End If
        lblDesignPicNum.Caption = (lblDesignPicNum.Caption - 1)
        Selection.Data5 = Val(lblDesignPicNum.Caption)
        If FileExist(DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption + 1) & GFX_EXT, False) Then
            cmdDesignNext.Enabled = True
        Else
            cmdDesignNext.Enabled = False
        End If
        If FileExist(DESIGN_PATH & Selection.Background & "\" & (lblDesignPicNum.Caption - 1) & GFX_EXT, False) Then
            cmdDesignPrevious.Enabled = True
        Else
            cmdDesignPrevious.Enabled = False
        End If
        Changed = True
        cmdDesignApply.Enabled = True
    End If
End Sub

Private Sub cmdDesignPreview_Click()
Dim n As Long

    If Selected = True And Changed = True Then
        ' Change selection values
        Selection.Data3 = Int(Val(txtDesignHeight.Text))
        Selection.Data4 = Int(Val(txtDesignWidth.Text))
        If Selection.Control = 0 And Selection.Background > 0 Then
            ' Show preview
            picDesignBackground(Selection.Background).Height = Selection.Data3
            picDesignBackground(Selection.Background).Width = Selection.Data4
            Call UpdateBackgroundSelection(Selection.Background)
        Else
            n = Selection.Control
            Select Case Selection.Background
                Case 1
                    lblDesignMenu(n).Height = Selection.Data3
                    lblDesignMenu(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 2
                    lblDesignLogin(n).Height = Selection.Data3
                    lblDesignLogin(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 3
                    lblDesignNewAcc(n).Height = Selection.Data3
                    lblDesignNewAcc(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 4
                    lblDesignDelAcc(n).Height = Selection.Data3
                    lblDesignDelAcc(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 5
                    lblDesignCredits(n).Height = Selection.Data3
                    lblDesignCredits(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 6
                    lblDesignChars(n).Height = Selection.Data3
                    lblDesignChars(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
                    
                Case 7
                    lblDesignNewChar(n).Height = Selection.Data3
                    lblDesignNewChar(n).Width = Selection.Data4
                    Call UpdateControlSelection(n, Selection.Background)
            End Select
        End If
        cmdDesignPreview.Enabled = False
        cmdDesignApply.Enabled = True
    End If
End Sub

Private Sub cmdDesignApply_Click()
Dim i As Long
Dim n As Long

    If Selected = True And Changed = True Then
        i = Selection.Background
        n = Selection.Control
        If Selection.Control = 0 And Selection.Background > 0 Then
            'Apply changes to background
            GUI(EditorIndex).Background(i).Data1 = Selection.Data1
            GUI(EditorIndex).Background(i).Data2 = Selection.Data2
            GUI(EditorIndex).Background(i).Data3 = Selection.Data3
            GUI(EditorIndex).Background(i).Data4 = Selection.Data4
            GUI(EditorIndex).Background(i).Data5 = Selection.Data5
        Else
            Select Case Selection.Background
                Case 1
                    GUI(EditorIndex).Menu(n).Data1 = Selection.Data1
                    GUI(EditorIndex).Menu(n).Data2 = Selection.Data2
                    GUI(EditorIndex).Menu(n).Data3 = Selection.Data3
                    GUI(EditorIndex).Menu(n).Data4 = Selection.Data4
                    
                Case 2
                    GUI(EditorIndex).Login(n).Data1 = Selection.Data1
                    GUI(EditorIndex).Login(n).Data2 = Selection.Data2
                    GUI(EditorIndex).Login(n).Data3 = Selection.Data3
                    GUI(EditorIndex).Login(n).Data4 = Selection.Data4
                    
                Case 3
                    GUI(EditorIndex).NewAcc(n).Data1 = Selection.Data1
                    GUI(EditorIndex).NewAcc(n).Data2 = Selection.Data2
                    GUI(EditorIndex).NewAcc(n).Data3 = Selection.Data3
                    GUI(EditorIndex).NewAcc(n).Data4 = Selection.Data4
                    
                Case 4
                    GUI(EditorIndex).DelAcc(n).Data1 = Selection.Data1
                    GUI(EditorIndex).DelAcc(n).Data2 = Selection.Data2
                    GUI(EditorIndex).DelAcc(n).Data3 = Selection.Data3
                    GUI(EditorIndex).DelAcc(n).Data4 = Selection.Data4
                    
                Case 5
                    GUI(EditorIndex).Credits(n).Data1 = Selection.Data1
                    GUI(EditorIndex).Credits(n).Data2 = Selection.Data2
                    GUI(EditorIndex).Credits(n).Data3 = Selection.Data3
                    GUI(EditorIndex).Credits(n).Data4 = Selection.Data4
                    
                Case 6
                    GUI(EditorIndex).Chars(n).Data1 = Selection.Data1
                    GUI(EditorIndex).Chars(n).Data2 = Selection.Data2
                    GUI(EditorIndex).Chars(n).Data3 = Selection.Data3
                    GUI(EditorIndex).Chars(n).Data4 = Selection.Data4
                    
                Case 7
                    GUI(EditorIndex).NewChar(n).Data1 = Selection.Data1
                    GUI(EditorIndex).NewChar(n).Data2 = Selection.Data2
                    GUI(EditorIndex).NewChar(n).Data3 = Selection.Data3
                    GUI(EditorIndex).NewChar(n).Data4 = Selection.Data4
            End Select
        End If
        Changed = False
        Call ResetSelection
    End If
End Sub

Private Sub optDesignOption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
    
    If Index = 1 Then
        For i = 1 To 5
            lblDesignMenu(i).Visible = True
        Next i
    Else
        For i = 1 To 5
            lblDesignMenu(i).Visible = False
        Next i
    End If
    For i = 2 To 7
        If Index = i Then
            picDesignBackground(i).Visible = True
        Else
            picDesignBackground(i).Visible = False
        End If
    Next i
End Sub

Private Sub txtDesignHeight_KeyPress(KeyAscii As Integer)
    cmdDesignPreview.Enabled = True
    Changed = True
    If (KeyAscii = vbKeyReturn) Then Call cmdDesignPreview_Click
End Sub

Private Sub txtDesignWidth_KeyPress(KeyAscii As Integer)
    cmdDesignPreview.Enabled = True
    Changed = True
    If (KeyAscii = vbKeyReturn) Then Call cmdDesignPreview_Click
End Sub

Private Sub cmdDesignSend_Click()
    Call GUIEditorSend
End Sub

Private Sub cmdDesignCancel_Click()
    Call GUIEditorCancel
End Sub

' **************************
' ** Selection processing **
' **************************

Private Sub ResetSelection()
Dim i As Long

    If Selected = True And Changed = False Then
        For i = 0 To 7
            shpDesignSelect(i).Visible = False
        Next i
        Selection.Background = 0
        Selection.Control = 0
        Selection.Data1 = 0
        Selection.Data2 = 0
        Selection.Data3 = 0
        Selection.Data4 = 0
        Selection.Data5 = 0
        fraDesignSelected.Visible = False
        If txtDesignHeight.Enabled = False Then txtDesignHeight.Enabled = True
        If txtDesignWidth.Enabled = False Then txtDesignWidth.Enabled = True
        txtDesignHeight.Text = "0"
        txtDesignWidth.Text = "0"
        lblDesignPicNum.Caption = "0"
        cmdDesignPrevious.Enabled = False
        cmdDesignNext.Enabled = False
        cmdDesignPreview.Enabled = False
        Selected = False
    End If
End Sub

Private Sub ResetChanges()
Dim i As Long
Dim n As Long

    i = Selection.Background
    n = Selection.Control
    If Selected = True And Changed = True Then
        If n = 0 And i > 0 Then
            picDesignBackground(i).Left = GUI(EditorIndex).Background(i).Data1
            picDesignBackground(i).Top = GUI(EditorIndex).Background(i).Data2
            picDesignBackground(i).Height = GUI(EditorIndex).Background(i).Data3
            picDesignBackground(i).Width = GUI(EditorIndex).Background(i).Data4
            If GUI(EditorIndex).Background(i).Data5 > 0 Then
                picDesignBackground(i).Picture = LoadPicture(App.Path & DESIGN_PATH & i & "\" & GUI(EditorIndex).Background(i).Data5 & GFX_EXT)
            Else
                picDesignBackground(i).Picture = LoadPicture()
            End If
        Else
            Select Case i
                Case 1
                    lblDesignMenu(n).Left = GUI(EditorIndex).Menu(n).Data1
                    lblDesignMenu(n).Top = GUI(EditorIndex).Menu(n).Data2
                    lblDesignMenu(n).Height = GUI(EditorIndex).Menu(n).Data3
                    lblDesignMenu(n).Width = GUI(EditorIndex).Menu(n).Data4
                    
                Case 2
                    lblDesignLogin(n).Left = GUI(EditorIndex).Login(n).Data1
                    lblDesignLogin(n).Top = GUI(EditorIndex).Login(n).Data2
                    lblDesignLogin(n).Height = GUI(EditorIndex).Login(n).Data3
                    lblDesignLogin(n).Width = GUI(EditorIndex).Login(n).Data4
                    
                Case 3
                    lblDesignNewAcc(n).Left = GUI(EditorIndex).NewAcc(n).Data1
                    lblDesignNewAcc(n).Top = GUI(EditorIndex).NewAcc(n).Data2
                    lblDesignNewAcc(n).Height = GUI(EditorIndex).NewAcc(n).Data3
                    lblDesignNewAcc(n).Width = GUI(EditorIndex).NewAcc(n).Data4
                    
                Case 4
                    lblDesignDelAcc(n).Left = GUI(EditorIndex).DelAcc(n).Data1
                    lblDesignDelAcc(n).Top = GUI(EditorIndex).DelAcc(n).Data2
                    lblDesignDelAcc(n).Height = GUI(EditorIndex).DelAcc(n).Data3
                    lblDesignDelAcc(n).Width = GUI(EditorIndex).DelAcc(n).Data4
                    
                Case 5
                    lblDesignCredits(n).Left = GUI(EditorIndex).Credits(n).Data1
                    lblDesignCredits(n).Top = GUI(EditorIndex).Credits(n).Data2
                    lblDesignCredits(n).Height = GUI(EditorIndex).Credits(n).Data3
                    lblDesignCredits(n).Width = GUI(EditorIndex).Credits(n).Data4
                    
                Case 6
                    lblDesignChars(n).Left = GUI(EditorIndex).Chars(n).Data1
                    lblDesignChars(n).Top = GUI(EditorIndex).Chars(n).Data2
                    lblDesignChars(n).Height = GUI(EditorIndex).Chars(n).Data3
                    lblDesignChars(n).Width = GUI(EditorIndex).Chars(n).Data4
                    
                Case 7
                    lblDesignNewChar(n).Left = GUI(EditorIndex).NewChar(n).Data1
                    lblDesignNewChar(n).Top = GUI(EditorIndex).NewChar(n).Data2
                    lblDesignNewChar(n).Height = GUI(EditorIndex).NewChar(n).Data3
                    lblDesignNewChar(n).Width = GUI(EditorIndex).NewChar(n).Data4
            End Select
        End If
        Changed = False
        cmdDesignApply.Enabled = False
    End If
End Sub

Private Sub UpdateBackgroundSelection(ByVal Background As Byte)
    If Background > 0 Then
        ' Use a different selection highlight if main menu background selected
        If Background = 1 Then
            fraDesignSelected.Visible = True
            shpDesignSelect(0).Width = picDesignBackground(Background).Width + 3
            shpDesignSelect(0).Height = picDesignBackground(Background).Height + 3
            shpDesignSelect(0).Top = picDesignBackground(Background).Top - 1
            shpDesignSelect(0).Left = picDesignBackground(Background).Left - 1
            shpDesignSelect(0).Visible = True
            txtDesignHeight.Text = picDesignBackground(Background).Height
            txtDesignWidth.Text = picDesignBackground(Background).Width
            fraDesignSelected.Caption = "Menu Background"
            lblDesignPicNum.Caption = STR(Selection.Data5)
            If FileExist(DESIGN_PATH & Background & "\" & (lblDesignPicNum.Caption + 1) & GFX_EXT, False) Then
                cmdDesignNext.Enabled = True
            Else
                cmdDesignNext.Enabled = False
            End If
            If FileExist(DESIGN_PATH & Background & "\" & (lblDesignPicNum.Caption - 1) & GFX_EXT, False) Then
                cmdDesignPrevious.Enabled = True
            Else
                cmdDesignPrevious.Enabled = False
            End If
        Else
            fraDesignSelected.Visible = True
            shpDesignSelect(1).Width = picDesignBackground(Background).Width + 3
            shpDesignSelect(1).Height = picDesignBackground(Background).Height + 3
            shpDesignSelect(1).Top = picDesignBackground(Background).Top - 1
            shpDesignSelect(1).Left = picDesignBackground(Background).Left - 1
            shpDesignSelect(1).Visible = True
            txtDesignHeight.Text = picDesignBackground(Background).Height
            txtDesignWidth.Text = picDesignBackground(Background).Width
            Select Case Background
                Case 2
                    fraDesignSelected.Caption = "Login Background"
                    
                Case 3
                    fraDesignSelected.Caption = "New Account Background"
                    
                Case 4
                    fraDesignSelected.Caption = "Delete Account Background"
                    
                Case 5
                    fraDesignSelected.Caption = "Credits Background"
                    
                Case 6
                    fraDesignSelected.Caption = "Select Character Background"
                    
                Case 7
                    fraDesignSelected.Caption = "New Character Background"
            End Select
            lblDesignPicNum.Caption = STR(Selection.Data5)
            If FileExist(DESIGN_PATH & Background & "\" & (lblDesignPicNum.Caption + 1) & GFX_EXT, False) Then
                cmdDesignNext.Enabled = True
            Else
                cmdDesignNext.Enabled = False
            End If
            If FileExist(DESIGN_PATH & Background & "\" & (lblDesignPicNum.Caption - 1) & GFX_EXT, False) Then
                cmdDesignPrevious.Enabled = True
            Else
                cmdDesignPrevious.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub UpdateControlSelection(ByVal Controll As Byte, ByVal Background As Byte)
    fraDesignSelected.Visible = True
    Select Case Background
        Case 1
            fraDesignSelected.Caption = "Menu Component"
            shpDesignSelect(Background).Width = lblDesignMenu(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignMenu(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignMenu(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignMenu(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignMenu(Controll).Height
            txtDesignWidth.Text = lblDesignMenu(Controll).Width
        Case 2
            fraDesignSelected.Caption = "Login Component"
            shpDesignSelect(Background).Width = lblDesignLogin(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignLogin(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignLogin(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignLogin(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignLogin(Controll).Height
            txtDesignWidth.Text = lblDesignLogin(Controll).Width
                    
        Case 3
            fraDesignSelected.Caption = "New Account Component"
            shpDesignSelect(Background).Width = lblDesignNewAcc(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignNewAcc(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignNewAcc(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignNewAcc(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignNewAcc(Controll).Height
            txtDesignWidth.Text = lblDesignNewAcc(Controll).Width
                    
        Case 4
            fraDesignSelected.Caption = "Delete Account Component"
            shpDesignSelect(Background).Width = lblDesignDelAcc(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignDelAcc(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignDelAcc(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignDelAcc(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignDelAcc(Controll).Height
            txtDesignWidth.Text = lblDesignDelAcc(Controll).Width
                    
        Case 5
            fraDesignSelected.Caption = "Credits Component"
            shpDesignSelect(Background).Width = lblDesignCredits(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignCredits(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignCredits(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignCredits(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignCredits(Controll).Height
            txtDesignWidth.Text = lblDesignCredits(Controll).Width
                    
        Case 6
            fraDesignSelected.Caption = "Select Character Component"
            shpDesignSelect(Background).Width = lblDesignChars(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignChars(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignChars(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignChars(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignChars(Controll).Height
            txtDesignWidth.Text = lblDesignChars(Controll).Width
                    
        Case 7
            fraDesignSelected.Caption = "New Character Component"
            shpDesignSelect(Background).Width = lblDesignNewChar(Controll).Width + 3
            shpDesignSelect(Background).Height = lblDesignNewChar(Controll).Height + 3
            shpDesignSelect(Background).Top = lblDesignNewChar(Controll).Top - 1
            shpDesignSelect(Background).Left = lblDesignNewChar(Controll).Left - 1
            shpDesignSelect(Background).Visible = True
            txtDesignHeight.Text = lblDesignNewChar(Controll).Height
            txtDesignWidth.Text = lblDesignNewChar(Controll).Width
    End Select
End Sub

Private Sub picDesignBackground_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Changed = False Then
            Call ResetSelection
        Else
            Call ResetChanges
            Call ResetSelection
        End If
    Else
        If Index = 1 And optDesignOption(1).Value = False Then Exit Sub
        If Selected = False Then
            Selection.Background = Index
            Selection.Control = 0
            Selection.Data1 = GUI(EditorIndex).Background(Index).Data1
            Selection.Data2 = GUI(EditorIndex).Background(Index).Data2
            Selection.Data3 = GUI(EditorIndex).Background(Index).Data3
            Selection.Data4 = GUI(EditorIndex).Background(Index).Data4
            Selection.Data5 = GUI(EditorIndex).Background(Index).Data5
            Selected = True
            Call UpdateBackgroundSelection(Selection.Background)
        End If
        If Index <> 1 And Selection.Control = 0 Then
            MouseXOffset = x
            MouseYOffset = y
            picDesignBackground(Index).Drag vbBeginDrag
        End If
    End If
End Sub

Private Sub picDesignBackground_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control = 0 Then
        ' Check if we're dropping the background onto itself
        If Index > 1 Then
            Source.Move x - MouseXOffset + Source.Left, y - MouseYOffset + Source.Top
        Else
            Source.Move x - MouseXOffset, y - MouseYOffset
        End If
        shpDesignSelect(1).Top = Source.Top - 1
        shpDesignSelect(1).Left = Source.Left - 1
        Selection.Data1 = picDesignBackground(Selection.Background).Left
        Selection.Data2 = picDesignBackground(Selection.Background).Top
    Else
        ' Check we're dropping the control on the right background
        If Selection.Background <> Index Then Exit Sub
        
        Source.Move x - MouseXOffset, y - MouseYOffset
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Select Case Selection.Background
            Case 1
                Selection.Data1 = lblDesignMenu(Selection.Control).Left
                Selection.Data2 = lblDesignMenu(Selection.Control).Top
                
            Case 2
                Selection.Data1 = lblDesignLogin(Selection.Control).Left
                Selection.Data2 = lblDesignLogin(Selection.Control).Top
                
            Case 3
                Selection.Data1 = lblDesignNewAcc(Selection.Control).Left
                Selection.Data2 = lblDesignNewAcc(Selection.Control).Top
                
            Case 4
                Selection.Data1 = lblDesignDelAcc(Selection.Control).Left
                Selection.Data2 = lblDesignDelAcc(Selection.Control).Top
                
            Case 5
                Selection.Data1 = lblDesignCredits(Selection.Control).Left
                Selection.Data2 = lblDesignCredits(Selection.Control).Top
                
            Case 6
                Selection.Data1 = lblDesignChars(Selection.Control).Left
                Selection.Data2 = lblDesignChars(Selection.Control).Top
                
            Case 7
                Selection.Data1 = lblDesignNewChar(Selection.Control).Left
                Selection.Data2 = lblDesignNewChar(Selection.Control).Top
        End Select
    End If
    Source.Visible = True
    cmdDesignApply.Enabled = True
    Changed = True
End Sub

Private Sub lblDesignMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 1
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).Menu(Index).Data1
        Selection.Data2 = GUI(EditorIndex).Menu(Index).Data2
        Selection.Data3 = GUI(EditorIndex).Menu(Index).Data3
        Selection.Data4 = GUI(EditorIndex).Menu(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignMenu(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignMenu_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Source.Move Int(x / 15) - MouseXOffset + lblDesignMenu(Index).Left, Int(y / 15) - MouseYOffset + lblDesignMenu(Index).Top
    shpDesignSelect(Selection.Background).Top = Source.Top - 1
    shpDesignSelect(Selection.Background).Left = Source.Left - 1
    Source.Visible = True
    cmdDesignApply.Enabled = True
    Changed = True
End Sub

Private Sub lblDesignLogin_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 2
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).Login(Index).Data1
        Selection.Data2 = GUI(EditorIndex).Login(Index).Data2
        Selection.Data3 = GUI(EditorIndex).Login(Index).Data3
        Selection.Data4 = GUI(EditorIndex).Login(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignLogin(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignLogin_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignLogin(Index).Left, Int(y / 15) - MouseYOffset + lblDesignLogin(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub

Private Sub lblDesignNewAcc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 3
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).NewAcc(Index).Data1
        Selection.Data2 = GUI(EditorIndex).NewAcc(Index).Data2
        Selection.Data3 = GUI(EditorIndex).NewAcc(Index).Data3
        Selection.Data4 = GUI(EditorIndex).NewAcc(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignNewAcc(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignNewAcc_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignNewAcc(Index).Left, Int(y / 15) - MouseYOffset + lblDesignNewAcc(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub

Private Sub lblDesignDelAcc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 4
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).DelAcc(Index).Data1
        Selection.Data2 = GUI(EditorIndex).DelAcc(Index).Data2
        Selection.Data3 = GUI(EditorIndex).DelAcc(Index).Data3
        Selection.Data4 = GUI(EditorIndex).DelAcc(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignDelAcc(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignDelAcc_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignDelAcc(Index).Left, Int(y / 15) - MouseYOffset + lblDesignDelAcc(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub

Private Sub lblDesignCredits_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 5
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).Credits(Index).Data1
        Selection.Data2 = GUI(EditorIndex).Credits(Index).Data2
        Selection.Data3 = GUI(EditorIndex).Credits(Index).Data3
        Selection.Data4 = GUI(EditorIndex).Credits(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignCredits(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignCredits_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignCredits(Index).Left, Int(y / 15) - MouseYOffset + lblDesignCredits(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub

Private Sub lblDesignChars_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 6
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).Chars(Index).Data1
        Selection.Data2 = GUI(EditorIndex).Chars(Index).Data2
        Selection.Data3 = GUI(EditorIndex).Chars(Index).Data3
        Selection.Data4 = GUI(EditorIndex).Chars(Index).Data4
        Selection.Data5 = 0
        Selected = True
        Call UpdateControlSelection(Selection.Control, Selection.Background)
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignChars(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignChars_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignChars(Index).Left, Int(y / 15) - MouseYOffset + lblDesignChars(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub

Private Sub lblDesignNewChar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = False Then
        Selection.Background = 7
        Selection.Control = Index
        Selection.Data1 = GUI(EditorIndex).NewChar(Index).Data1
        Selection.Data2 = GUI(EditorIndex).NewChar(Index).Data2
        Selection.Data3 = GUI(EditorIndex).NewChar(Index).Data3
        Selection.Data4 = GUI(EditorIndex).NewChar(Index).Data4
        Selection.Data5 = 0
        Selected = True
        If Index = 9 Or Index = 10 Or Index = 13 Or Index = 14 Then
            Call UpdateControlSelection(Selection.Control, Selection.Background)
        Else
            txtDesignHeight.Enabled = False
            txtDesignWidth.Enabled = False
            Call UpdateControlSelection(Selection.Control, Selection.Background)
        End If
    End If
    If Selection.Control = Index Then
        MouseXOffset = Int(x / 15)
        MouseYOffset = Int(y / 15)
        lblDesignNewChar(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lblDesignNewChar_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If Selection.Control > 0 Then
        Source.Move Int(x / 15) - MouseXOffset + lblDesignNewChar(Index).Left, Int(y / 15) - MouseYOffset + lblDesignNewChar(Index).Top
        shpDesignSelect(Selection.Background).Top = Source.Top - 1
        shpDesignSelect(Selection.Background).Left = Source.Left - 1
        Source.Visible = True
        cmdDesignApply.Enabled = True
        Changed = True
    End If
End Sub
