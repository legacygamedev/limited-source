VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCClient 
   BorderStyle     =   0  'None
   Caption         =   "CClientDev"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   766
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1022
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox picSetRSpawn 
         Height          =   495
         Left            =   10920
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   149
         TabIndex        =   278
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Label lblSetRSpawn 
            Alignment       =   2  'Center
            Caption         =   "LeftClick to set resource spawn - RightClick to cancel."
            Height          =   450
            Left            =   30
            TabIndex        =   279
            Top             =   15
            Width           =   2175
         End
      End
      Begin VB.PictureBox picCoord 
         Height          =   855
         Left            =   12960
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         Begin VB.Label lblCoordY 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   480
            TabIndex        =   97
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lblCoordX 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   480
            TabIndex        =   96
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label45 
            Caption         =   "Y :"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label44 
            Caption         =   "X :"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.PictureBox picMapEditor 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7815
         Left            =   0
         ScaleHeight     =   519
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   513
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   7725
         Begin VB.PictureBox picMapProperties 
            Height          =   7455
            Left            =   120
            ScaleHeight     =   493
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   493
            TabIndex        =   44
            Top             =   120
            Visible         =   0   'False
            Width           =   7455
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   19
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   277
               Top             =   6120
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   18
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   276
               Top             =   5760
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   17
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   275
               Top             =   5400
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   16
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   274
               Top             =   5040
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   15
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   273
               Top             =   4680
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   14
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   272
               Top             =   4320
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   13
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   271
               Top             =   3960
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   12
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   270
               Top             =   5760
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   11
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   269
               Top             =   5400
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   10
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   268
               Top             =   5040
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   9
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   267
               Top             =   4680
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   8
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   266
               Top             =   4320
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   7
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   265
               Top             =   3960
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   6
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   264
               Top             =   6120
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   5
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   263
               Top             =   5760
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   4
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   262
               Top             =   5400
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   3
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   261
               Top             =   5040
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   2
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   260
               Top             =   4680
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   1
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   259
               Top             =   4320
               Width           =   375
            End
            Begin VB.CheckBox chkSetRSpawn 
               Caption         =   "x/y"
               Height          =   300
               Index           =   0
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   258
               Top             =   3960
               Width           =   375
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   10
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   238
               Top             =   5040
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   11
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   237
               Top             =   5400
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   12
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   236
               Top             =   5760
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   13
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   235
               Top             =   3960
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   14
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   234
               Top             =   4320
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   15
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   233
               Top             =   4680
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   16
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   232
               Top             =   5040
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   17
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   231
               Top             =   5400
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   18
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   230
               Top             =   5760
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   19
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   229
               Top             =   6120
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   9
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   4680
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   8
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   4320
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   7
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   3960
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   6
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   6120
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   5
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   5760
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   4
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   132
               Top             =   5400
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   3
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   131
               Top             =   5040
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   2
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   130
               Top             =   4680
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   1
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   4320
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesResource 
               Height          =   315
               Index           =   0
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   3960
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesNpc 
               Height          =   315
               Index           =   4
               Left            =   3960
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   3000
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesNpc 
               Height          =   315
               Index           =   3
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   3000
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesNpc 
               Height          =   315
               Index           =   2
               Left            =   5040
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   2520
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesNpc 
               Height          =   315
               Index           =   1
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   99
               Top             =   2520
               Width           =   1815
            End
            Begin VB.ComboBox cmbPropertiesNpc 
               Height          =   315
               Index           =   0
               ItemData        =   "frmCClient.frx":0000
               Left            =   720
               List            =   "frmCClient.frx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   2520
               Width           =   1815
            End
            Begin VB.ComboBox cmbMoral 
               Height          =   315
               ItemData        =   "frmCClient.frx":0004
               Left            =   1320
               List            =   "frmCClient.frx":000E
               TabIndex        =   56
               Text            =   "cmbMoral"
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdPropertiesOk 
               Caption         =   "OK"
               Height          =   495
               Left            =   480
               TabIndex        =   54
               Top             =   6720
               Width           =   4095
            End
            Begin VB.TextBox txtBootY 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5760
               TabIndex        =   53
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtBootX 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5760
               TabIndex        =   52
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtBootMap 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5760
               TabIndex        =   51
               Top             =   120
               Width           =   735
            End
            Begin VB.TextBox txtRight 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3120
               TabIndex        =   50
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtLeft 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3120
               TabIndex        =   49
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtDown 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1320
               TabIndex        =   48
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtUp 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1320
               TabIndex        =   47
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtMapName 
               Height          =   285
               Left            =   1320
               TabIndex        =   46
               Top             =   120
               Width           =   2535
            End
            Begin VB.CommandButton cmdPropertiesCancel 
               Caption         =   "Cancel"
               Height          =   495
               Left            =   4800
               TabIndex        =   45
               Top             =   6720
               Width           =   1695
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               Caption         =   "Resources"
               Height          =   255
               Left            =   3000
               TabIndex        =   250
               Top             =   3600
               Width           =   1695
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               Caption         =   "NPC's"
               Height          =   255
               Left            =   3480
               TabIndex        =   249
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Music"
               Height          =   255
               Left            =   4440
               TabIndex        =   248
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label22 
               Caption         =   "Boot Y"
               Height          =   255
               Left            =   4800
               TabIndex        =   247
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label21 
               Caption         =   "Boot X"
               Height          =   255
               Left            =   4800
               TabIndex        =   246
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label20 
               Caption         =   "Boot Map"
               Height          =   255
               Left            =   4680
               TabIndex        =   245
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label19 
               Caption         =   "Moral"
               Height          =   255
               Left            =   480
               TabIndex        =   244
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label18 
               Caption         =   "Right"
               Height          =   255
               Left            =   2400
               TabIndex        =   243
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label17 
               Caption         =   "Left"
               Height          =   255
               Left            =   2400
               TabIndex        =   242
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label16 
               Caption         =   "Down"
               Height          =   255
               Left            =   600
               TabIndex        =   241
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label15 
               Caption         =   "Up"
               Height          =   255
               Left            =   600
               TabIndex        =   240
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label14 
               Caption         =   "Name"
               Height          =   255
               Left            =   480
               TabIndex        =   239
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.PictureBox picPushBlock 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   7395
            TabIndex        =   208
            Top             =   5160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.CommandButton cmdPushBlockCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   4080
               TabIndex        =   219
               Top             =   1920
               Width           =   2175
            End
            Begin VB.CommandButton cmdPushBlockOK 
               Caption         =   "OK"
               Height          =   375
               Left            =   1200
               TabIndex        =   212
               Top             =   1920
               Width           =   2175
            End
            Begin VB.Frame fraDir3 
               Caption         =   "Direction 3"
               Height          =   1575
               Left            =   5160
               TabIndex        =   211
               Top             =   120
               Width           =   1215
               Begin VB.OptionButton optDir3Right 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   228
                  Top             =   1200
                  Width           =   735
               End
               Begin VB.OptionButton optDir3Left 
                  Caption         =   "Left"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   227
                  Top             =   960
                  Width           =   735
               End
               Begin VB.OptionButton optDir3Down 
                  Caption         =   "Down"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   226
                  Top             =   720
                  Width           =   735
               End
               Begin VB.OptionButton optDir3Up 
                  Caption         =   "Up"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   225
                  Top             =   480
                  Width           =   735
               End
               Begin VB.OptionButton optDir3None 
                  Caption         =   "None"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   224
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin VB.Frame fraDir2 
               Caption         =   "Direction 2"
               Height          =   1575
               Left            =   3240
               TabIndex        =   210
               Top             =   120
               Width           =   1215
               Begin VB.OptionButton optDir2Right 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   223
                  Top             =   1200
                  Width           =   735
               End
               Begin VB.OptionButton optDir2Left 
                  Caption         =   "Left"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   222
                  Top             =   960
                  Width           =   735
               End
               Begin VB.OptionButton optDir2Down 
                  Caption         =   "Down"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   221
                  Top             =   720
                  Width           =   735
               End
               Begin VB.OptionButton optDir2Up 
                  Caption         =   "Up"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   220
                  Top             =   480
                  Width           =   735
               End
               Begin VB.OptionButton optDir2None 
                  Caption         =   "None"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   218
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.Frame fraDir1 
               Caption         =   "Direction 1"
               Height          =   1575
               Left            =   1080
               TabIndex        =   209
               Top             =   120
               Width           =   1215
               Begin VB.OptionButton optDir1Right 
                  Caption         =   "Right"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   217
                  Top             =   1200
                  Width           =   735
               End
               Begin VB.OptionButton optDir1Left 
                  Caption         =   "Left"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   216
                  Top             =   960
                  Width           =   735
               End
               Begin VB.OptionButton optDir1Down 
                  Caption         =   "Down"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   215
                  Top             =   720
                  Width           =   735
               End
               Begin VB.OptionButton optDir1Up 
                  Caption         =   "Up"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   214
                  Top             =   480
                  Width           =   615
               End
               Begin VB.OptionButton optDir1None 
                  Caption         =   "None"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   213
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
         End
         Begin VB.PictureBox picKeyOpen 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   7395
            TabIndex        =   84
            Top             =   5160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.HScrollBar scrlKeyOpenY 
               Height          =   255
               Left            =   1920
               TabIndex        =   91
               Top             =   960
               Width           =   3735
            End
            Begin VB.HScrollBar scrlKeyOpenX 
               Height          =   255
               Left            =   1920
               TabIndex        =   89
               Top             =   600
               Width           =   3735
            End
            Begin VB.CommandButton cmdKeyOpenCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3960
               TabIndex        =   86
               Top             =   1800
               Width           =   2295
            End
            Begin VB.CommandButton cmdKeyOpenOk 
               Caption         =   "OK"
               Height          =   375
               Left            =   1080
               TabIndex        =   85
               Top             =   1800
               Width           =   2415
            End
            Begin VB.Label lblKeyOpenY 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   5880
               TabIndex        =   92
               Top             =   960
               Width           =   90
            End
            Begin VB.Label lblKeyOpenX 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   5880
               TabIndex        =   90
               Top             =   600
               Width           =   90
            End
            Begin VB.Label Label43 
               Caption         =   "Y"
               Height          =   255
               Left            =   1440
               TabIndex        =   88
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label42 
               Caption         =   "X"
               Height          =   255
               Left            =   1440
               TabIndex        =   87
               Top             =   600
               Width           =   255
            End
         End
         Begin VB.PictureBox picMapKey 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   7395
            TabIndex        =   75
            Top             =   5160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.CheckBox chkMapKeyTake 
               Caption         =   "Take key away upon use (not operational)"
               Height          =   255
               Left            =   2640
               TabIndex        =   83
               Top             =   1200
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.HScrollBar scrlMapKeyItem 
               Height          =   255
               Left            =   2160
               Max             =   1000
               Min             =   1
               TabIndex        =   81
               Top             =   600
               Value           =   1
               Width           =   3255
            End
            Begin VB.CommandButton cmdMapKeyCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   4080
               TabIndex        =   79
               Top             =   1800
               Width           =   2655
            End
            Begin VB.CommandButton cmdMapKeyOk 
               Caption         =   "OK"
               Height          =   375
               Left            =   840
               TabIndex        =   78
               Top             =   1800
               Width           =   2535
            End
            Begin VB.Label lblMapKeyItem 
               AutoSize        =   -1  'True
               Caption         =   "1"
               Height          =   195
               Left            =   5640
               TabIndex        =   82
               Top             =   600
               Width           =   90
            End
            Begin VB.Label Label41 
               Caption         =   "Item No."
               Height          =   255
               Left            =   1200
               TabIndex        =   80
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lblMapKeyName 
               Height          =   255
               Left            =   2160
               TabIndex        =   77
               Top             =   120
               Width           =   3735
            End
            Begin VB.Label Label40 
               Caption         =   "Item"
               Height          =   255
               Left            =   1440
               TabIndex        =   76
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.PictureBox picMapWarp 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   7395
            TabIndex        =   33
            Top             =   5160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.HScrollBar scrlMapWarpX 
               Height          =   255
               Left            =   2400
               Max             =   31
               TabIndex        =   38
               Top             =   720
               Width           =   3255
            End
            Begin VB.HScrollBar scrlMapWarpY 
               Height          =   255
               Left            =   2400
               Max             =   23
               TabIndex        =   37
               Top             =   1080
               Width           =   3255
            End
            Begin VB.CommandButton cmdMapWarpOK 
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1320
               TabIndex        =   36
               Top             =   1680
               Width           =   2055
            End
            Begin VB.CommandButton cmdMapWarpCancel 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4200
               TabIndex        =   35
               Top             =   1680
               Width           =   2055
            End
            Begin VB.TextBox txtWarpMap 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1080
               TabIndex        =   34
               Top             =   120
               Width           =   3855
            End
            Begin VB.Label Label1 
               Caption         =   "Map"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   43
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   42
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "Y"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   41
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label lblMapWarpX 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5520
               TabIndex        =   40
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblMapWarpY 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5520
               TabIndex        =   39
               Top             =   1080
               Width           =   495
            End
         End
         Begin VB.PictureBox picMapItem 
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   7395
            TabIndex        =   63
            Top             =   5160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.CommandButton cmdMapItemCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   4080
               TabIndex        =   72
               Top             =   1800
               Width           =   2415
            End
            Begin VB.CommandButton cmdMapItemOk 
               Caption         =   "OK"
               Height          =   375
               Left            =   960
               TabIndex        =   71
               Top             =   1800
               Width           =   2655
            End
            Begin VB.HScrollBar scrlMapItem 
               Height          =   255
               Left            =   2520
               Max             =   1000
               Min             =   1
               TabIndex        =   69
               Top             =   960
               Value           =   1
               Width           =   3375
            End
            Begin VB.HScrollBar scrlMapItemValue 
               Height          =   255
               Left            =   2520
               Max             =   1000
               TabIndex        =   68
               Top             =   1320
               Width           =   3375
            End
            Begin VB.Label lblMapItemValue 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   6240
               TabIndex        =   74
               Top             =   1320
               Width           =   90
            End
            Begin VB.Label lblMapItem 
               AutoSize        =   -1  'True
               Caption         =   "1"
               Height          =   195
               Left            =   6240
               TabIndex        =   70
               Top             =   960
               Width           =   90
            End
            Begin VB.Label Label39 
               Caption         =   "Value"
               Height          =   255
               Left            =   1320
               TabIndex        =   67
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label38 
               Caption         =   "Item No."
               Height          =   255
               Left            =   1320
               TabIndex        =   66
               Top             =   960
               Width           =   855
            End
            Begin VB.Label lblMapItemName 
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
               Left            =   2160
               TabIndex        =   65
               Top             =   240
               Width           =   4575
            End
            Begin VB.Label Label37 
               Caption         =   "Item"
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
               Left            =   840
               TabIndex        =   64
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.OptionButton optBuildLayer 
            Caption         =   "Build Layer"
            Height          =   375
            Left            =   3240
            TabIndex        =   163
            Top             =   6600
            Width           =   1335
         End
         Begin VB.CommandButton cmdFlushDirection 
            Caption         =   "Flush Blocking"
            Height          =   255
            Left            =   3600
            TabIndex        =   58
            Top             =   7440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optDirectionView 
            Caption         =   "Directional Blocking"
            Height          =   375
            Left            =   3240
            TabIndex        =   57
            Top             =   6960
            Width           =   1335
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   375
            Left            =   5520
            TabIndex        =   55
            Top             =   6720
            Width           =   1575
         End
         Begin VB.Frame fraAttribs 
            Caption         =   "Attributes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   360
            TabIndex        =   24
            Top             =   5160
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton optPushBlock 
               Caption         =   "Push Block"
               Height          =   255
               Left            =   120
               TabIndex        =   207
               Top             =   840
               Width           =   1215
            End
            Begin VB.OptionButton optKey 
               Caption         =   "Key"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   31
               Top             =   1200
               Width           =   1215
            End
            Begin VB.OptionButton optNpcAvoid 
               Caption         =   "Npc Avoid"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton optItem 
               Caption         =   "Item"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1440
               TabIndex        =   29
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdClear2 
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   28
               Top             =   1920
               Width           =   1215
            End
            Begin VB.OptionButton optWarp 
               Caption         =   "Warp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1440
               TabIndex        =   27
               Top             =   960
               Width           =   735
            End
            Begin VB.OptionButton optBlocked 
               Caption         =   "Blocked"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optKeyOpen 
               Caption         =   "Key Open"
               Height          =   240
               Left            =   120
               TabIndex        =   25
               Top             =   1440
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdProperties 
            Caption         =   "Properties"
            Height          =   495
            Left            =   4800
            TabIndex        =   32
            Top             =   5640
            Width           =   2055
         End
         Begin VB.Frame fraLayers 
            Caption         =   "Layers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   360
            TabIndex        =   11
            Top             =   5160
            Width           =   2535
            Begin VB.CommandButton cmdFill 
               Caption         =   "Fill"
               Height          =   255
               Left            =   1680
               TabIndex        =   73
               Top             =   1800
               Width           =   615
            End
            Begin VB.OptionButton optLight 
               Caption         =   "Light Layer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1560
               Width           =   1455
            End
            Begin VB.OptionButton optFringe2 
               Caption         =   "Fringe 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1320
               Width           =   1215
            End
            Begin VB.OptionButton optFAnim 
               Caption         =   "Animation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   21
               Top             =   1080
               Width           =   1215
            End
            Begin VB.OptionButton optMask2 
               Caption         =   "Mask 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   840
               Width           =   1215
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   16
               Top             =   1920
               Width           =   1215
            End
            Begin VB.OptionButton optFringe 
               Caption         =   "Fringe"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   1080
               Width           =   1215
            End
            Begin VB.OptionButton optAnim 
               Caption         =   "Animation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   14
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optMask 
               Caption         =   "Mask"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optGround 
               Caption         =   "Ground"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5520
            TabIndex        =   19
            Top             =   7200
            Width           =   1575
         End
         Begin VB.OptionButton optAttribs 
            Caption         =   "Attributes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   18
            Top             =   6360
            Width           =   1575
         End
         Begin VB.OptionButton optLayers 
            Caption         =   "Layers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   6000
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.PictureBox picSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   3600
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   5280
            Width           =   480
         End
         Begin VB.VScrollBar scrlPicture 
            Height          =   4815
            Left            =   7080
            Max             =   255
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox picBack 
            BackColor       =   &H00000000&
            Height          =   4860
            Left            =   300
            ScaleHeight     =   320
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   448
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   6780
            Begin VB.PictureBox picBackSelect 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   1215
               Left            =   0
               ScaleHeight     =   81
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   89
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   0
               Width           =   1335
            End
         End
      End
      Begin VB.PictureBox picRightClickMenu 
         Height          =   4335
         Left            =   8040
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdDevChat 
            Caption         =   "Chat"
            Height          =   375
            Left            =   120
            TabIndex        =   257
            Top             =   3360
            Width           =   1935
         End
         Begin VB.CommandButton cmdDesignTool 
            Caption         =   "Menu Design"
            Height          =   375
            Left            =   120
            TabIndex        =   251
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditQuest 
            Caption         =   "Edit Quest"
            Height          =   375
            Left            =   120
            TabIndex        =   191
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditShop 
            Caption         =   "Edit Shop"
            Height          =   375
            Left            =   120
            TabIndex        =   122
            Top             =   2040
            Width           =   1935
         End
         Begin VB.CommandButton cmdQuit 
            Caption         =   "Quit"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   3840
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditMap 
            Caption         =   "Edit Map"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditSkill 
            Caption         =   "Edit Skill"
            Height          =   375
            Left            =   120
            TabIndex        =   154
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditSpell 
            Caption         =   "Edit Spell"
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditNpc 
            Caption         =   "Edit NPC"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton cmdEditItem 
            Caption         =   "Edit Item"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.PictureBox picDevChat 
         Height          =   2175
         Left            =   240
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   481
         TabIndex        =   252
         Top             =   9120
         Visible         =   0   'False
         Width           =   7275
         Begin VB.CheckBox chkDevChatPin 
            Appearance      =   0  'Flat
            Caption         =   "Pin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   256
            Top             =   90
            Width           =   615
         End
         Begin RichTextLib.RichTextBox txtChat 
            Height          =   1335
            Left            =   120
            TabIndex        =   254
            TabStop         =   0   'False
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2355
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmCClient.frx":0023
         End
         Begin VB.Timer tmrDevChat 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   5640
            Top             =   120
         End
         Begin VB.Label lblDevChatCancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   6900
            TabIndex        =   255
            Top             =   90
            Width           =   165
         End
         Begin VB.Label lblChat 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   253
            Top             =   1800
            Width           =   6975
         End
      End
      Begin VB.PictureBox picIndex 
         Height          =   3135
         Left            =   5640
         ScaleHeight     =   3075
         ScaleWidth      =   4395
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lstIndex 
            Height          =   2400
            ItemData        =   "frmCClient.frx":00A5
            Left            =   120
            List            =   "frmCClient.frx":00A7
            TabIndex        =   60
            Top             =   120
            Width           =   4095
         End
         Begin VB.CommandButton cmdIndexCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2280
            TabIndex        =   62
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CommandButton cmdIndexOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   2640
            Width           =   2055
         End
      End
      Begin VB.PictureBox picEditSpell 
         Height          =   4575
         Left            =   5880
         ScaleHeight     =   301
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   103
         Top             =   4080
         Visible         =   0   'False
         Width           =   3975
         Begin VB.HScrollBar scrlSpellPic 
            Height          =   255
            Left            =   720
            Max             =   50
            TabIndex        =   124
            Top             =   1200
            Width           =   1695
         End
         Begin VB.PictureBox picSpellPic 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   3000
            ScaleHeight     =   36
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   36
            TabIndex        =   123
            Top             =   1080
            Width           =   540
         End
         Begin VB.CommandButton cmdEditSpellCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2040
            TabIndex        =   121
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditSpellOk 
            Caption         =   "OK"
            Height          =   375
            Left            =   360
            TabIndex        =   120
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Frame fraSpellGiveItem 
            Caption         =   "Give Item"
            Height          =   1215
            Left            =   240
            TabIndex        =   113
            Top             =   2760
            Visible         =   0   'False
            Width           =   3495
            Begin VB.HScrollBar scrlSpellItemValue 
               Height          =   255
               Left            =   720
               Max             =   255
               TabIndex        =   117
               Top             =   720
               Width           =   2295
            End
            Begin VB.HScrollBar scrlSpellItemnum 
               Height          =   255
               Left            =   720
               Max             =   255
               TabIndex        =   116
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label lblSpellItemValue 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3120
               TabIndex        =   119
               Top             =   720
               Width           =   90
            End
            Begin VB.Label lblSpellItemNum 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3120
               TabIndex        =   118
               Top             =   360
               Width           =   90
            End
            Begin VB.Label Label53 
               Caption         =   "Value"
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label52 
               Caption         =   "Item"
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame fraSpellStats 
            Caption         =   "Vitals Data"
            Height          =   1215
            Left            =   240
            TabIndex        =   112
            Top             =   2760
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txtSpellStatMod 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2160
               TabIndex        =   166
               Text            =   "1"
               Top             =   780
               Width           =   855
            End
            Begin VB.ComboBox cmbSpellStat 
               Height          =   315
               ItemData        =   "frmCClient.frx":00A9
               Left            =   360
               List            =   "frmCClient.frx":00C2
               Style           =   2  'Dropdown List
               TabIndex        =   164
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label12 
               Caption         =   "Stat Mod per Level"
               Height          =   255
               Left            =   480
               TabIndex        =   165
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.ComboBox cmbSpellType 
            Height          =   315
            ItemData        =   "frmCClient.frx":00FC
            Left            =   240
            List            =   "frmCClient.frx":0109
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   2280
            Width           =   3375
         End
         Begin VB.HScrollBar scrlSpellLevelReq 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   109
            Top             =   1800
            Width           =   2535
         End
         Begin VB.ComboBox cmbSpellClassReq 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtSpellName 
            Height          =   285
            Left            =   840
            TabIndex        =   106
            Top             =   120
            Width           =   2775
         End
         Begin VB.PictureBox picSpellsBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   41
            TabIndex        =   127
            Top             =   4200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Timer tmrSpellPic 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   1800
            Top             =   4200
         End
         Begin VB.Label Label46 
            Caption         =   "Pic"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblSpellPic 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2520
            TabIndex        =   125
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label lblSpellLevelReq 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3480
            TabIndex        =   110
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label Label50 
            Caption         =   "Level"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label49 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.PictureBox picEditSkill 
         Height          =   4575
         Left            =   5880
         ScaleHeight     =   301
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   138
         Top             =   3720
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Frame fraSkillChance 
            Caption         =   "Chance Mod"
            Height          =   1215
            Left            =   240
            TabIndex        =   158
            Top             =   2640
            Width           =   3495
            Begin VB.TextBox txtSkillChanceGain 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2400
               TabIndex        =   161
               Text            =   "1"
               Top             =   770
               Width           =   495
            End
            Begin VB.ComboBox cmbSkillChance 
               Height          =   315
               ItemData        =   "frmCClient.frx":0126
               Left            =   360
               List            =   "frmCClient.frx":0139
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label9 
               Caption         =   "%"
               Height          =   255
               Left            =   3000
               TabIndex        =   162
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label8 
               Caption         =   "Percentage gain per Level"
               Height          =   255
               Left            =   360
               TabIndex        =   160
               Top             =   840
               Width           =   1935
            End
         End
         Begin VB.Frame fraSkillVitals 
            Caption         =   "Vital Mod"
            Height          =   1215
            Left            =   240
            TabIndex        =   153
            Top             =   2640
            Visible         =   0   'False
            Width           =   3495
            Begin VB.ComboBox cmbWepType 
               Height          =   315
               ItemData        =   "frmCClient.frx":0180
               Left            =   1680
               List            =   "frmCClient.frx":019C
               Style           =   2  'Dropdown List
               TabIndex        =   167
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox txtSkillAttributeGain 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2160
               TabIndex        =   157
               Text            =   "1"
               Top             =   770
               Width           =   615
            End
            Begin VB.ComboBox cmbSkillAttribute 
               Height          =   315
               ItemData        =   "frmCClient.frx":01D2
               Left            =   360
               List            =   "frmCClient.frx":01E8
               Style           =   2  'Dropdown List
               TabIndex        =   155
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Points per Level gained"
               Height          =   255
               Left            =   360
               TabIndex        =   156
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdEditSkillCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2040
            TabIndex        =   149
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditSkillOk 
            Caption         =   "OK"
            Height          =   375
            Left            =   360
            TabIndex        =   148
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Timer tmrSkillPic 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   1800
            Top             =   4200
         End
         Begin VB.PictureBox picSkillsBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   615
            TabIndex        =   147
            Top             =   4200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmbSkillType 
            Height          =   315
            ItemData        =   "frmCClient.frx":021E
            Left            =   240
            List            =   "frmCClient.frx":0231
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   2280
            Width           =   3375
         End
         Begin VB.HScrollBar scrlSkillLevelReq 
            Height          =   255
            Left            =   840
            TabIndex        =   144
            Top             =   1800
            Width           =   2535
         End
         Begin VB.HScrollBar scrlSkillPic 
            Height          =   255
            Left            =   720
            TabIndex        =   142
            Top             =   1200
            Width           =   1695
         End
         Begin VB.PictureBox picSkillPic 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3000
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   141
            Top             =   1080
            Width           =   495
         End
         Begin VB.ComboBox cmbSkillClassReq 
            Height          =   315
            ItemData        =   "frmCClient.frx":025E
            Left            =   240
            List            =   "frmCClient.frx":0260
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtSkillName 
            Height          =   285
            Left            =   840
            TabIndex        =   139
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "Level"
            Height          =   255
            Left            =   240
            TabIndex        =   152
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Pic"
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblSkillLevelReq 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3480
            TabIndex        =   145
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label lblSkillPic 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2520
            TabIndex        =   143
            Top             =   1200
            Width           =   90
         End
      End
      Begin VB.PictureBox picEditQuest 
         Height          =   6855
         Left            =   5640
         ScaleHeight     =   6795
         ScaleWidth      =   4275
         TabIndex        =   168
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
         Begin VB.Frame fraTradeQuest 
            Caption         =   "Trade Quest"
            Height          =   1215
            Left            =   240
            TabIndex        =   200
            Top             =   2400
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.Frame fraFetchQuest 
            Caption         =   "Fetch Quest"
            Height          =   1215
            Left            =   240
            TabIndex        =   199
            Top             =   2400
            Visible         =   0   'False
            Width           =   3855
            Begin VB.HScrollBar scrlQuestFetchQuantity 
               Height          =   255
               Left            =   1080
               Max             =   1000
               TabIndex        =   205
               Top             =   720
               Width           =   2295
            End
            Begin VB.HScrollBar scrlQuestFetchItem 
               Height          =   255
               Left            =   1080
               Max             =   50
               TabIndex        =   203
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label lblQuestFetchQuantity 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3480
               TabIndex        =   206
               Top             =   720
               Width           =   90
            End
            Begin VB.Label lblQuestFetchItem 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3480
               TabIndex        =   204
               Top             =   360
               Width           =   90
            End
            Begin VB.Label Label13 
               Caption         =   "Quantity"
               Height          =   255
               Left            =   360
               TabIndex        =   202
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label11 
               Caption         =   "Item"
               Height          =   255
               Left            =   360
               TabIndex        =   201
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.CommandButton cmdQuestCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   2160
            TabIndex        =   193
            Top             =   6240
            Width           =   1815
         End
         Begin VB.CommandButton cmdQuestOK 
            Caption         =   "OK"
            Height          =   495
            Left            =   240
            TabIndex        =   192
            Top             =   6240
            Width           =   1815
         End
         Begin VB.Frame fraKillQuest 
            Caption         =   "Kill Quest"
            Height          =   1215
            Left            =   240
            TabIndex        =   189
            Top             =   2400
            Visible         =   0   'False
            Width           =   3855
            Begin VB.HScrollBar scrlQuestKillQuantity 
               Height          =   255
               Left            =   1080
               Max             =   100
               TabIndex        =   197
               Top             =   720
               Width           =   2295
            End
            Begin VB.HScrollBar scrlQuestKillNpc 
               Height          =   255
               Left            =   1080
               Max             =   50
               TabIndex        =   194
               Top             =   360
               Value           =   1
               Width           =   2295
            End
            Begin VB.Label lblQuestKillQuantity 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3480
               TabIndex        =   198
               Top             =   720
               Width           =   90
            End
            Begin VB.Label Label10 
               Caption         =   "Quantity"
               Height          =   255
               Left            =   240
               TabIndex        =   196
               Top             =   720
               Width           =   615
            End
            Begin VB.Label lblQuestKillNpc 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   3480
               TabIndex        =   195
               Top             =   360
               Width           =   90
            End
            Begin VB.Label Label32 
               Caption         =   "Kill Npc"
               Height          =   255
               Left            =   240
               TabIndex        =   190
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.TextBox txtQuestRewardVal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   188
            Text            =   "0"
            Top             =   4440
            Width           =   1455
         End
         Begin VB.HScrollBar scrlQuestReward 
            Height          =   255
            Left            =   360
            Max             =   50
            TabIndex        =   185
            Top             =   4080
            Value           =   1
            Width           =   3255
         End
         Begin VB.ComboBox cmbQuestType 
            Height          =   315
            ItemData        =   "frmCClient.frx":0262
            Left            =   1080
            List            =   "frmCClient.frx":026F
            Style           =   2  'Dropdown List
            TabIndex        =   182
            Top             =   600
            Width           =   3015
         End
         Begin VB.ComboBox cmbQuestClass 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   1080
            Width           =   2775
         End
         Begin VB.HScrollBar scrlQuestLevelMax 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   178
            Top             =   1920
            Width           =   2295
         End
         Begin VB.HScrollBar scrlQuestLevelMin 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   174
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txtQuestDescription 
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   172
            Top             =   5040
            Width           =   3975
         End
         Begin VB.TextBox txtQuestName 
            Height          =   375
            Left            =   1080
            TabIndex        =   170
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label txtQuestReward 
            Height          =   255
            Left            =   1200
            TabIndex        =   184
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label31 
            Caption         =   "Reward Value :"
            Height          =   255
            Left            =   960
            TabIndex        =   187
            Top             =   4440
            Width           =   1215
         End
         Begin VB.Label lblQuestReward 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3720
            TabIndex        =   186
            Top             =   4080
            Width           =   90
         End
         Begin VB.Label Label30 
            Caption         =   "Reward :"
            Height          =   255
            Left            =   360
            TabIndex        =   183
            Top             =   3720
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "Type"
            Height          =   255
            Left            =   360
            TabIndex        =   181
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblQuestLevelMax 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3720
            TabIndex        =   179
            Top             =   1920
            Width           =   90
         End
         Begin VB.Label Label28 
            Caption         =   "Level Maximum"
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Level Minimum"
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblQuestLevelMin 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3720
            TabIndex        =   175
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label Label26 
            Caption         =   "Class Required"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   171
            Top             =   4800
            Width           =   1095
         End
         Begin VB.Label Label24 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   169
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   14760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCClient"
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

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not InItemsEditor And Not InNpcEditor And Not InSpellEditor And Not InSkillEditor And Not InShopEditor Then
        If Button = 2 Then
            If InEditor Then
                If frmCClient.picMapEditor.Visible = True Then
                    frmCClient.picMapEditor.Visible = False
                Else
                    If frmCClient.picSetRSpawn.Visible = True Then
                        frmCClient.chkSetRSpawn(RSpawnNum).Value = Unchecked
                        RSpawnNum = 20
                        frmCClient.picSetRSpawn.Visible = False
                        frmCClient.picCoord.Visible = False
                    End If
                    frmCClient.picMapEditor.Visible = True
                End If
            Else
                If Int(x / 32) > (MAX_MAPX - 4) Or Int(y / 32) > (MAX_MAPY - 5) Then
                    frmCClient.picRightClickMenu.Top = y - 150
                    frmCClient.picRightClickMenu.Left = x - 120
                    frmCClient.picRightClickMenu.Visible = True
                Else
                    frmCClient.picRightClickMenu.Top = y - 20
                    frmCClient.picRightClickMenu.Left = x - 20
                    frmCClient.picRightClickMenu.Visible = True
                End If
            End If
        Else
            If Button = 1 Then
                frmCClient.picRightClickMenu.Visible = False
                Call EditorMouseDown(Button, Shift, x, y)
                'Call PlayerSearch(Button, Shift, x, y)
            End If
        End If
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Integer

    x1 = Int(x / PIC_X)
    If x1 < 0 Or x1 > MAX_MAPX Then
        Exit Sub
    End If
    y1 = Int(y / PIC_Y)
    If y1 < 0 Or y1 > MAX_MAPY Then
        Exit Sub
    End If
    
    If InEditor Then
        If picSetRSpawn.Visible = False Then
            If ((Map.Tile(x1, y1).Type <> TILE_TYPE_WALKABLE) And (Map.Tile(x1, y1).Type <> TILE_TYPE_BLOCKED) And (Map.Tile(x1, y1).Type <> TILE_TYPE_NPCAVOID)) Or ((GetPlayerX(MyIndex) = x1) And (GetPlayerY(MyIndex) = y1)) Then
                lblCoordX.Caption = x1
                lblCoordY.Caption = y1
                picCoord.Top = y + 8
                picCoord.Left = x + 8
                picCoord.Visible = True
            Else
                picCoord.Visible = False
            End If
        Else
            picSetRSpawn.Left = x + 8
            picSetRSpawn.Top = y - picSetRSpawn.Height
            lblCoordX.Caption = x1
            lblCoordY.Caption = y1
            picCoord.Top = y + 8
            picCoord.Left = x + 8
        End If
    End If
    Call EditorMouseDown(Button, Shift, x, y)
End Sub

Private Sub picScreen_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub picScreen_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x - MouseXOffset, y - MouseYOffset
    picScreen.SetFocus
    Source.Visible = True
End Sub

Private Sub picScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    picRightClickMenu.Visible = False
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub picScreen_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckInput(0, KeyCode, Shift)
End Sub

' ********************
' * Right Click Menu *
' ********************

Private Sub cmdEditMap_Click()
    Call SendRequestEditMap
End Sub

Private Sub cmdEditItem_Click()
    Call SendRequestEditItem
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdEditNpc_Click()
    Call SendRequestEditNpc
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdEditSpell_Click()
    Call SendRequestEditSpell
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdEditSkill_Click()
    Call SendRequestEditSkill
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdEditShop_Click()
    Call SendRequestEditShop
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdEditQuest_Click()
    Call SendRequestEditQuest
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdDesignTool_Click()
    Call SendRequestEditGUI
    picRightClickMenu.Visible = False
End Sub

Private Sub cmdDevChat_Click()
    picRightClickMenu.Visible = False
    chkDevChatPin.Value = Unchecked
    picDevChat.Visible = True
    tmrDevChat.Enabled = True
End Sub

Private Sub cmdQuit_Click()
    Call GameDestroy
End Sub

' ********************
' * Map Editor Stuff *
' ********************

Private Sub optLayers_Click()
    If optLayers.Value = True Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
        cmdFlushDirection.Visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
        cmdFlushDirection.Visible = False
    End If
End Sub

Private Sub optDirectionView_Click()
    If optDirectionView.Value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = False
        cmdFlushDirection.Visible = True
    End If
End Sub

Private Sub optBuildLayer_Click()
    If optBuildLayer.Value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = False
        cmdFlushDirection.Visible = False
    End If
End Sub

Private Sub cmdFlushDirection_Click()
Dim x As Long, y As Long
   
    ' Go through each tile and set each direction to walkable
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
             With Map.Tile(x, y)
                 .WalkUp = 0
                 .WalkDown = 0
                 .WalkLeft = 0
                 .WalkRight = 0
             End With
        Next y
    Next x
End Sub

Private Sub cmdFill_Click()
Dim y As Long
Dim x As Long
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If frmCClient.optLayers.Value = True Then
                With Map.Tile(x, y)
                If frmCClient.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                If frmCClient.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                If frmCClient.optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                If frmCClient.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                If frmCClient.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                If frmCClient.optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                If frmCClient.optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                End With
            End If
        Next x
    Next y
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        picMapEditor.Visible = False
        Exit Sub
    End If
    Call EditorChooseTile(Button, Shift, x, y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call EditorChooseTile(Button, Shift, x, y)
End Sub

Private Sub cmdSend_Click()
    picCoord.Visible = False
    GettingMap = True
    Call EditorSend
End Sub

Private Sub cmdCancel_Click()
    picCoord.Visible = False
    Call EditorCancel
End Sub

Private Sub cmdProperties_Click()
    Call SetMapProperties
    picMapProperties.Visible = True
End Sub

Private Sub optWarp_Click()
    picMapWarp.Visible = True
End Sub

Private Sub optItem_Click()
    scrlMapItem.Max = MAX_ITEMS
    lblMapItemName.Caption = Trim(Item(scrlMapItem.Value).Name)
    picMapItem.Visible = True
End Sub

Private Sub optKey_Click()
    lblMapKeyName.Caption = Trim(Item(scrlMapKeyItem.Value).Name)
    scrlMapKeyItem.Max = MAX_ITEMS
    picMapKey.Visible = True
End Sub

Private Sub optKeyOpen_Click()
    scrlKeyOpenX.Max = MAX_MAPX
    scrlKeyOpenY.Max = MAX_MAPY
    picKeyOpen.Visible = True
End Sub

Private Sub optPushBlock_Click()
    picPushBlock.Visible = True
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub

Private Sub scrlPicture_Scroll()
    Call EditorTileScroll
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

Private Sub picMapEditor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picMapEditor.Drag vbBeginDrag
End Sub

Private Sub picMapEditor_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picMapEditor.Left - MouseXOffset, y + picMapEditor.Top - MouseYOffset
    picScreen.SetFocus
    Source.Visible = True
End Sub

Private Sub picBackSelect_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picMapEditor.Left + picBackSelect.Left - MouseXOffset, y + picMapEditor.Top + picBackSelect.Top - MouseYOffset
    picScreen.SetFocus
    Source.Visible = True
End Sub

' ****************
' * pic Map Warp *
' ****************

Private Sub cmdMapWarpOk_Click()
    EditorWarpMap = Val(txtWarpMap.Text)
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picMapWarp.Visible = False
End Sub

Private Sub cmdMapWarpCancel_Click()
    picMapWarp.Visible = False
End Sub

Private Sub scrlMapWarpX_Change()
    lblMapWarpX.Caption = STR(scrlMapWarpX.Value)
End Sub

Private Sub scrlMapWarpY_Change()
    lblMapWarpY.Caption = STR(scrlMapWarpY.Value)
End Sub

Private Sub txtWarpMap_Change()
    If Val(txtWarpMap.Text) <= 0 Or Val(txtWarpMap.Text) > MAX_MAPS Then
        txtWarpMap.Text = ""
    End If
End Sub

' **********************
' * pic Map Properties *
' **********************

Private Sub SetMapProperties()
Dim x As Long, y As Long, i As Long

    txtMapName.Text = Trim(Map.Name)
    txtUp.Text = STR(Map.Up)
    txtDown.Text = STR(Map.Down)
    txtLeft.Text = STR(Map.Left)
    txtRight.Text = STR(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    'scrlMusic.Value = Map.Music
    txtBootMap.Text = STR(Map.BootMap)
    txtBootX.Text = STR(Map.BootX)
    txtBootY.Text = STR(Map.BootY)
    
    For x = 1 To MAX_MAP_NPCS
        cmbPropertiesNpc(x - 1).Clear
        cmbPropertiesNpc(x - 1).AddItem "No Npc"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_NPCS
            If Npc(y).Behavior <> NPC_BEHAVIOR_RESOURCE Then
                cmbPropertiesNpc(x - 1).AddItem y & ": " & Trim(Npc(y).Name)
            Else
                cmbPropertiesNpc(x - 1).AddItem "**Not NPC type**"
            End If
        Next x
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        cmbPropertiesNpc(i - 1).ListIndex = Map.Npc(i)
    Next i
    
    For x = 1 To MAX_MAP_RESOURCES
        cmbPropertiesResource(x - 1).Clear
        cmbPropertiesResource(x - 1).AddItem "No Resource"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_RESOURCES
            If Npc(y).Behavior = NPC_BEHAVIOR_RESOURCE Then
                cmbPropertiesResource(x - 1).AddItem y & ": " & Trim(Npc(y).Name)
            Else
                cmbPropertiesResource(x - 1).AddItem "**Not RESOURCE**"
            End If
        Next x
    Next y
    
    For i = 1 To MAX_MAP_RESOURCES
        cmbPropertiesResource(i - 1).ListIndex = Map.Resource(i)
    Next i
    
    For i = 1 To MAX_MAP_RESOURCES
        If Map.RSpawn(i).RSx > 0 Or Map.RSpawn(i).RSy > 0 Then
            chkSetRSpawn(i - 1).Value = Checked
        Else
            chkSetRSpawn(i - 1).Value = Unchecked
        End If
    Next i
End Sub

'Private Sub scrlMusic_Change()
    'lblMusic.Caption = STR(scrlMusic.Value)
'End Sub

Private Sub cmdPropertiesOk_Click()
Dim x As Long, y As Long, i As Long

    Map.Name = txtMapName.Text
    Map.Up = Val(txtUp.Text)
    Map.Down = Val(txtDown.Text)
    Map.Left = Val(txtLeft.Text)
    Map.Right = Val(txtRight.Text)
    Map.Moral = cmbMoral.ListIndex
    'Map.Music = scrlMusic.Value
    Map.BootMap = Val(txtBootMap.Text)
    Map.BootX = Val(txtBootX.Text)
    Map.BootY = Val(txtBootY.Text)
    
    For i = 1 To MAX_MAP_NPCS
        Map.Npc(i) = cmbPropertiesNpc(i - 1).ListIndex
    Next i
    
    For i = 1 To MAX_MAP_RESOURCES
        Map.Resource(i) = cmbPropertiesResource(i - 1).ListIndex
    Next i
    
    picMapProperties.Visible = False
End Sub

Private Sub chkSetRSpawn_Click(Index As Integer)
    If cmbPropertiesResource(Index).ListIndex > 0 Then
        RSpawnNum = Index
        picMapEditor.Visible = False
        picSetRSpawn.Visible = True
        picCoord.Visible = True
        picScreen.SetFocus
    End If
    chkSetRSpawn(Index).Value = Unchecked
End Sub

Private Sub cmdPropertiesCancel_Click()
    picMapProperties.Visible = False
End Sub

' ******************
' * pic Push Block *
' ******************

Private Sub cmdPushBlockOK_Click()
    If optDir1None.Value = True Then
        PushDir1 = 0
    ElseIf optDir1Up.Value = True Then
        PushDir1 = 1
    ElseIf optDir1Down.Value = True Then
        PushDir1 = 2
    ElseIf optDir1Left.Value = True Then
        PushDir1 = 3
    ElseIf optDir1Right.Value = True Then
        PushDir1 = 4
    Else
        PushDir1 = 0
    End If
    If optDir2None.Value = True Then
        PushDir2 = 0
    ElseIf optDir2Up.Value = True Then
        PushDir2 = 1
    ElseIf optDir2Down.Value = True Then
        PushDir2 = 2
    ElseIf optDir2Left.Value = True Then
        PushDir2 = 3
    ElseIf optDir2Right.Value = True Then
        PushDir2 = 4
    Else
        PushDir2 = 0
    End If
    If optDir3None.Value = True Then
        PushDir3 = 0
    ElseIf optDir3Up.Value = True Then
        PushDir3 = 1
    ElseIf optDir3Down.Value = True Then
        PushDir3 = 2
    ElseIf optDir3Left.Value = True Then
        PushDir3 = 3
    ElseIf optDir3Right.Value = True Then
        PushDir3 = 4
    Else
        PushDir3 = 0
    End If
    picPushBlock.Visible = False
End Sub

Private Sub cmdPushBlockCancel_Click()
    picPushBlock.Visible = False
End Sub

' ****************
' * Editor Index *
' ****************

Private Sub cmdIndexOk_Click()
    EditorIndex = lstIndex.ListIndex + 1
    
    If InItemsEditor = True Then
        Call SendData("EDITITEM" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InNpcEditor = True Then
        Call SendData("EDITNPC" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InShopEditor = True Then
        Call SendData("EDITSHOP" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InSpellEditor = True Then
        Call SendData("EDITSPELL" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InSkillEditor = True Then
        Call SendData("EDITSKILL" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InQuestEditor = True Then
        Call SendData("EDITQUEST" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    If InGUIEditor = True Then
        Call SendData("EDITGUI" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
    picIndex.Visible = False
End Sub

Private Sub cmdIndexCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InSpellEditor = False
    InSkillEditor = False
    InQuestEditor = False
    InGUIEditor = False
    picIndex.Visible = False
End Sub

' ************************
' * Spell Editor Section *
' ************************

Private Sub cmbSpellType_Click()
    If cmbSpellType.ListIndex = SPELL_TYPE_STAT Then
        fraSpellStats.Visible = True
        fraSpellGiveItem.Visible = False
    Else
        fraSpellStats.Visible = False
        fraSpellGiveItem.Visible = True
    End If
End Sub

Private Sub scrlSpellItemNum_Change()
    If scrlSpellItemnum <> 0 Then
        fraSpellGiveItem.Caption = "Give Item " & Trim(Item(scrlSpellItemnum.Value).Name)
    Else
        fraSpellGiveItem.Caption = "Give Item"
    End If
    lblSpellItemNum.Caption = STR(scrlSpellItemnum.Value)
End Sub

Private Sub scrlSpellPic_Change()
    lblSpellPic.Caption = STR(scrlSpellPic.Value)
End Sub

Private Sub scrlSpellItemValue_Change()
    lblSpellItemValue.Caption = STR(scrlSpellItemValue.Value)
End Sub

Private Sub scrlSpellLevelReq_Change()
    lblSpellLevelReq.Caption = STR(scrlSpellLevelReq.Value)
End Sub

'Private Sub scrlSpellVitalMod_Change()
    'lblSpellVitalMod.Caption = STR(scrlSpellVitalMod.Value)
'End Sub

Private Sub cmdEditSpellOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdEditSpellCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub tmrSpellPic_Timer()
    Call SpellEditorBltItem
End Sub

' ************************
' * Skill Editor Section *
' ************************

Private Sub cmbSkillType_Click()
    If cmbSkillType.ListIndex = SKILL_TYPE_ATTRIBUTE Then
        fraSkillVitals.Visible = True
        'Exit Sub
    Else
        fraSkillVitals.Visible = False
    End If
    If cmbSkillType.ListIndex = SKILL_TYPE_CHANCE Then
        fraSkillChance.Visible = True
        'Exit Sub
    Else
        fraSkillChance.Visible = False
    End If
End Sub

Private Sub cmbSkillAttribute_Click()
    If cmbSkillAttribute.ListIndex = SKILL_ATTRIBUTE_STR Then
        cmbWepType.Enabled = True
    Else
        cmbWepType.ListIndex = 0
        cmbWepType.Enabled = False
    End If
End Sub

Private Sub scrlSkillPic_Change()
    lblSkillPic.Caption = STR(scrlSkillPic.Value)
End Sub

Private Sub scrlSkillLevelReq_Change()
    lblSkillLevelReq.Caption = STR(scrlSkillLevelReq.Value)
End Sub

Private Sub cmdEditSkillOk_Click()
    Call SkillEditorOk
End Sub

Private Sub cmdEditSkillCancel_Click()
    Call SkillEditorCancel
End Sub

Private Sub tmrSkillPic_Timer()
    Call SkillEditorBltItem
End Sub

' ************************
' * Quest Editor Section *
' ************************

Private Sub cmbQuestType_Click()
    Select Case cmbQuestType.ListIndex
        Case QUEST_TYPE_KILL
            fraKillQuest.Visible = True
            fraFetchQuest.Visible = False
            fraTradeQuest.Visible = False
            
        Case QUEST_TYPE_FETCH
            fraKillQuest.Visible = False
            fraFetchQuest.Visible = True
            fraTradeQuest.Visible = False
            
        Case QUEST_TYPE_TRADE
            fraKillQuest.Visible = False
            fraFetchQuest.Visible = False
            fraTradeQuest.Visible = True
    End Select
End Sub

Private Sub scrlQuestLevelMin_Change()
    lblQuestLevelMin.Caption = STR(scrlQuestLevelMin.Value)
End Sub

Private Sub scrlQuestLevelMax_Change()
    lblQuestLevelMax.Caption = STR(scrlQuestLevelMax.Value)
End Sub

Private Sub scrlQuestReward_Change()
    lblQuestReward.Caption = STR(scrlQuestReward.Value)
    If scrlQuestReward.Value > 0 Then
        txtQuestReward.Caption = Trim(Item(scrlQuestReward.Value).Name)
    Else
        txtQuestReward.Caption = "None"
    End If
End Sub

Private Sub scrlQuestKillNpc_Change()
    lblQuestKillNpc.Caption = STR(scrlQuestKillNpc.Value)
    If scrlQuestKillNpc.Value > 0 Then
        If Npc(scrlQuestKillNpc.Value).Behavior = NPC_BEHAVIOR_FRIENDLY Then
            fraKillQuest.Caption = "Kill " & Trim(Npc(scrlQuestKillNpc.Value).Name) & " - FRIENDLY NPC"
        ElseIf Npc(scrlQuestKillNpc.Value).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
            fraKillQuest.Caption = "Kill " & Trim(Npc(scrlQuestKillNpc.Value).Name) & " - SHOPKEEPER NPC"
        Else
            fraKillQuest.Caption = "Kill " & Trim(Npc(scrlQuestKillNpc.Value).Name)
        End If
    Else
        fraKillQuest.Caption = "Kill Quest"
    End If
End Sub

Private Sub scrlQuestKillQuantity_Change()
    lblQuestKillQuantity.Caption = STR(scrlQuestKillQuantity.Value)
End Sub

Private Sub scrlQuestFetchItem_Change()
    lblQuestFetchItem.Caption = STR(scrlQuestFetchItem.Value)
    If scrlQuestFetchItem.Value > 0 Then
        fraFetchQuest.Caption = "Fetch " & Trim(Item(scrlQuestFetchItem.Value).Name)
    Else
        fraFetchQuest.Caption = "Fetch Quest"
    End If
End Sub

Private Sub scrlQuestFetchQuantity_Change()
    lblQuestFetchQuantity.Caption = STR(scrlQuestFetchQuantity.Value)
End Sub

Private Sub cmdQuestOK_Click()
    Call QuestEditorOk
End Sub

Private Sub cmdQuestCancel_Click()
    Call QuestEditorCancel
End Sub

' ********************
' * Dev Chat Section *
' ********************

Private Sub tmrDevChat_Timer()
    picDevChat.Visible = False
    tmrDevChat.Enabled = False
End Sub

Private Sub lblDevChatCancel_Click()
    picDevChat.Visible = False
    tmrDevChat.Enabled = False
    chkDevChatPin.Value = Unchecked
End Sub

Private Sub chkDevChatPin_Click()
    If chkDevChatPin.Value = Checked Then
        tmrDevChat.Enabled = False
        If picScreen.Visible Then picScreen.SetFocus
    Else
        tmrDevChat.Enabled = True
        picScreen.SetFocus
    End If
End Sub

Private Sub picDevChat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseXOffset = x
    MouseYOffset = y
    picDevChat.Drag vbBeginDrag
End Sub

Private Sub picDevChat_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x + picDevChat.Left - MouseXOffset, y + picDevChat.Top - MouseYOffset
    picScreen.SetFocus
End Sub

Private Sub txtChat_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move (x / 15) + picDevChat.Left + txtChat.Left - MouseXOffset, (y / 15) + picDevChat.Top + txtChat.Top - MouseYOffset
    picScreen.SetFocus
End Sub


' ********************
' * Map Item Section *
' ********************

Private Sub cmdMapItemOk_Click()
    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picMapItem.Visible = False
    picMapEditor.Visible = False
End Sub

Private Sub cmdMapItemCancel_Click()
    picMapItem.Visible = False
End Sub

Private Sub scrlMapItem_Change()
    lblMapItem.Caption = STR(scrlMapItem.Value)
    lblMapItemName.Caption = Trim(Item(scrlMapItem.Value).Name)
End Sub

Private Sub scrlMapItemValue_Change()
    lblMapItemValue.Caption = STR(scrlMapItemValue.Value)
End Sub

' *******************
' * Map Key Section *
' *******************

Private Sub cmdMapKeyOk_Click()
    KeyEditorNum = scrlMapKeyItem.Value
    KeyEditorTake = chkMapKeyTake.Value
    picMapKey.Visible = False
End Sub

Private Sub scrlMapKeyItem_Change()
    lblMapKeyItem.Caption = STR(scrlMapKeyItem.Value)
    lblMapKeyName.Caption = Trim(Item(scrlMapKeyItem.Value).Name)
End Sub

Private Sub cmdMapKeyCancel_Click()
    picMapKey.Visible = False
End Sub

' ********************
' * Key Open Section *
' ********************

Private Sub cmdKeyOpenOk_Click()
    KeyOpenEditorX = scrlKeyOpenX.Value
    KeyOpenEditorY = scrlKeyOpenY.Value
    picKeyOpen.Visible = False
End Sub

Private Sub cmdKeyOpenCancel_Click()
    picKeyOpen.Visible = False
End Sub

Private Sub scrlKeyOpenX_Change()
    lblKeyOpenX.Caption = STR(scrlKeyOpenX.Value)
End Sub

Private Sub scrlKeyOpenY_Change()
    lblKeyOpenY.Caption = STR(scrlKeyOpenY.Value)
End Sub
