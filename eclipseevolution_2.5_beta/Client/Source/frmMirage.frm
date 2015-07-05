VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{96366485-4AD2-4BC8-AFBF-B1FC132616A5}#2.0#0"; "VBMP.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Evolution"
   ClientHeight    =   9135
   ClientLeft      =   555
   ClientTop       =   780
   ClientWidth     =   13725
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMirage.frx":030A
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   915
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Mousetmr 
      Interval        =   1200
      Left            =   1200
      Top             =   240
   End
   Begin VB.PictureBox picrciinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   12000
      ScaleHeight     =   1425
      ScaleWidth      =   1665
      TabIndex        =   246
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Line Line13 
         X1              =   0
         X2              =   1680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblrcigift 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gift"
         Height          =   255
         Left            =   0
         TabIndex        =   251
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblrcidrop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drop"
         Height          =   255
         Left            =   0
         TabIndex        =   250
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   1680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblrciuse 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use/Equip"
         Height          =   255
         Left            =   0
         TabIndex        =   249
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   1680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblrciinfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         Height          =   255
         Left            =   0
         TabIndex        =   248
         Top             =   480
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   1680
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblrciname 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   0
         TabIndex        =   247
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox picInv3 
      Appearance      =   0  'Flat
      BackColor       =   &H00828B82&
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
      Height          =   3675
      Left            =   2160
      Picture         =   "frmMirage.frx":3C24
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   5985
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   19
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   245
         Top             =   1785
         Width           =   480
      End
      Begin VB.PictureBox ShoesImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4230
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   244
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox LegsImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4230
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   243
         Top             =   1755
         Width           =   495
      End
      Begin VB.PictureBox RingImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5280
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   242
         Top             =   1740
         Width           =   495
      End
      Begin VB.PictureBox ArmorImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4230
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   241
         Top             =   1230
         Width           =   495
      End
      Begin VB.PictureBox ShieldImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4755
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   240
         Top             =   1230
         Width           =   495
      End
      Begin VB.PictureBox NecklaceImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3180
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   239
         Top             =   1740
         Width           =   495
      End
      Begin VB.PictureBox WeaponImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3690
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   238
         Top             =   1215
         Width           =   495
      End
      Begin VB.PictureBox HelmetImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4230
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   237
         Top             =   705
         Width           =   495
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   29
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   236
         Top             =   2820
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   28
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   235
         Top             =   2820
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   27
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   234
         Top             =   2820
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   26
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   233
         Top             =   2820
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   25
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   232
         Top             =   2820
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   24
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   231
         Top             =   2295
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   23
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   230
         Top             =   2295
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   22
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   229
         Top             =   2295
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   21
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   228
         Top             =   2295
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   20
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   227
         Top             =   2295
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   18
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   226
         Top             =   1770
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   17
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   225
         Top             =   1770
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   16
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   224
         Top             =   1770
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   15
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   223
         Top             =   1770
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   14
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   222
         Top             =   1245
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   13
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   221
         Top             =   1245
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   12
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   220
         Top             =   1245
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   219
         Top             =   1245
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   218
         Top             =   1245
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   217
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   216
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   215
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   214
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   213
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   2310
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   212
         Top             =   195
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   1785
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   211
         Top             =   195
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1260
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   210
         Top             =   195
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   735
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   209
         Top             =   195
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   208
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   253
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   7
         Left            =   600
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   0
         Left            =   600
         Top             =   720
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   1
         Left            =   0
         Top             =   720
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   2
         Left            =   600
         Top             =   720
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   3
         Left            =   600
         Top             =   600
         Width           =   540
      End
      Begin VB.Shape SelectedItem 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   525
         Left            =   1200
         Top             =   240
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   540
      End
      Begin VB.Shape EquipS 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   540
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2040
         TabIndex        =   3
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   690
      End
   End
   Begin VB.PictureBox picrclick 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   11760
      ScaleHeight     =   1665
      ScaleWidth      =   1665
      TabIndex        =   200
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label lbrcname 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   0
         TabIndex        =   206
         Top             =   0
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lbrclvl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lvl 40 Mage"
         Height          =   255
         Left            =   0
         TabIndex        =   205
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1680
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbrctrade 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade"
         Height          =   255
         Left            =   0
         TabIndex        =   204
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   1680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lbrcchat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Chat"
         Height          =   255
         Left            =   0
         TabIndex        =   203
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   1680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lbrcparty 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         Height          =   255
         Left            =   0
         TabIndex        =   202
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   1680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lbrcpm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PM"
         Height          =   255
         Left            =   0
         TabIndex        =   201
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   1680
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   2160
      ScaleHeight     =   3345
      ScaleWidth      =   2400
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   2400
      Begin VB.TextBox txtAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   23
         Top             =   585
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   22
         Top             =   345
         Width           =   1575
      End
      Begin VB.CommandButton cmdTrainee 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         Caption         =   "Make Trainee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   975
         Width           =   1815
      End
      Begin VB.CommandButton cmdMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         Caption         =   "Make Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1305
         Width           =   1815
      End
      Begin VB.CommandButton cmdDisown 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         Caption         =   "Disown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1650
         Width           =   1815
      End
      Begin VB.CommandButton cmdAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         Caption         =   "Change Access"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   25
         Top             =   615
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   210
         TabIndex        =   24
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Index           =   0
      Left            =   2160
      ScaleHeight     =   3345
      ScaleWidth      =   2400
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label cmdLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Guild"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   32
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label lblRank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1425
         TabIndex        =   31
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lblGuild 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1425
         TabIndex        =   30
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Rank :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   570
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   480
         TabIndex        =   28
         Top             =   645
         Width           =   825
      End
   End
   Begin CodeSenseCtl.CodeSense CS 
      Height          =   3255
      Left            =   13680
      OleObjectBlob   =   "frmMirage.frx":4B8D6
      TabIndex        =   195
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox piccharstats 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   2160
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   98
      Top             =   3120
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Skills"
         Height          =   255
         Left            =   0
         TabIndex        =   166
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Key config"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label AddDef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   117
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label AddStr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   116
         Top             =   1920
         Width           =   165
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   115
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   114
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label AddMagi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   113
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label AddSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   112
         Top             =   1560
         Width           =   165
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   111
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   110
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   109
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label lblSTATWINDOW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CHARACTER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   840
         TabIndex        =   107
         Top             =   2640
         Width           =   1050
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL :  "
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   106
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   105
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   104
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   103
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   102
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   101
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENERGY :  "
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   100
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   99
         Top             =   840
         Width           =   1050
      End
   End
   Begin VB.PictureBox picWhosOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   9360
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         ItemData        =   "frmMirage.frx":4BA3C
         Left            =   45
         List            =   "frmMirage.frx":4BA3E
         TabIndex        =   15
         Top             =   120
         Width           =   2310
      End
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   2160
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2835
         ItemData        =   "frmMirage.frx":4BA40
         Left            =   45
         List            =   "frmMirage.frx":4BA42
         TabIndex        =   5
         Top             =   60
         Width           =   2310
      End
      Begin VB.Label lblForgetSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   35
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblCast 
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   3120
         Width           =   375
      End
   End
   Begin VB.Timer tmrGameClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13320
      Top             =   8760
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00828B82&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   8160
      ScaleHeight     =   287
      ScaleMode       =   0  'User
      ScaleWidth      =   175
      TabIndex        =   40
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   252
         Top             =   0
         Width           =   255
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   360
         TabIndex        =   51
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requirements-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   50
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   49
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   48
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Add-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   46
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   45
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   44
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magi: XXXXX Speed: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   41
         Top             =   2040
         Width           =   2655
      End
   End
   Begin VBMP.VBMPlayer BGSPlayer 
      Height          =   1095
      Left            =   0
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VBMP.VBMPlayer SoundPlayer 
      Height          =   1095
      Left            =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   18
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   97
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   17
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   96
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   16
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   95
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   14
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   93
      Top             =   240
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmote 
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   15
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   94
         Top             =   240
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   13
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   92
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   12
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   91
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   11
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   90
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   10
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   89
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   9
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   88
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   8
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   87
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   7
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   86
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   6
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   85
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   84
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   4
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   83
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   2
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   82
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   3
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   81
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   80
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   79
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picEmote 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   19
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   78
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   4905
      Left            =   9120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   52
      Top             =   2040
      Visible         =   0   'False
      Width           =   2625
      Begin VB.HScrollBar ScrlResolution 
         Height          =   255
         Left            =   120
         Max             =   3
         Min             =   1
         TabIndex        =   77
         Top             =   4200
         Value           =   1
         Width           =   2295
      End
      Begin VB.CheckBox chkplayerbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mini HP Bar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkplayername 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   62
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkbubblebar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Speech Bubbles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   61
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show NPC HP Bars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkplayerdamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Damage Above Head"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcdamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Damage Above Heads"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkmusic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   57
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chksound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   56
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   4
         TabIndex        =   55
         Top             =   3720
         Value           =   6
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   4560
         Width           =   2325
      End
      Begin VB.CheckBox chkAutoScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto Scroll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-Screen Resolution-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblLines 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Screen Text Line Amount: 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   69
         Top             =   3480
         Width           =   2325
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Player Data-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-NPC Data-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   67
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Sound/Music Data-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   66
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Chat Data-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   65
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.PictureBox picUber 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   2130
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   39
      Top             =   390
      Width           =   9600
      Begin VB.PictureBox Skills 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   3480
         ScaleHeight     =   4065
         ScaleWidth      =   5025
         TabIndex        =   147
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   2
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   186
            Top             =   840
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   1
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   187
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   1
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   188
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   6
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   183
            Top             =   3240
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   5
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   184
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   5
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   185
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   5
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   180
            Top             =   2640
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   4
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   181
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   4
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   182
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   4
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   177
            Top             =   2040
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   3
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   178
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   3
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   179
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   3
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   174
            Top             =   1440
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   2
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   175
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   2
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   176
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   1
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   171
            Top             =   240
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   0
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   172
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Index           =   0
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   173
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.CommandButton exit 
            Caption         =   "x"
            Height          =   255
            Left            =   4750
            TabIndex        =   170
            Top             =   25
            Width           =   255
         End
         Begin VB.CommandButton next 
            Caption         =   ">"
            Height          =   255
            Left            =   4725
            TabIndex        =   169
            Top             =   3840
            Width           =   255
         End
         Begin VB.CommandButton back 
            Caption         =   "<"
            Height          =   255
            Left            =   50
            TabIndex        =   168
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label34 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exp:   Level:"
            Height          =   255
            Left            =   3720
            TabIndex        =   191
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label33 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name:"
            Height          =   255
            Left            =   840
            TabIndex        =   190
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label32 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Icon:"
            Height          =   255
            Left            =   120
            TabIndex        =   189
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   5
            Left            =   3720
            TabIndex        =   165
            Top             =   3240
            Width           =   405
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   4
            Left            =   3720
            TabIndex        =   164
            Top             =   2640
            Width           =   405
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   163
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   162
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   161
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Exp 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   160
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   159
            Top             =   3240
            Width           =   405
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   158
            Top             =   2640
            Width           =   405
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   157
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   156
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   155
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Level 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   154
            Top             =   240
            Width           =   405
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   153
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   152
            Top             =   2640
            Width           =   2775
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   151
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   150
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   149
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label skillname 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   148
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.PictureBox BoxKeys 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   2760
         ScaleHeight     =   2025
         ScaleWidth      =   6825
         TabIndex        =   118
         Top             =   5160
         Visible         =   0   'False
         Width           =   6855
         Begin VB.ComboBox ItemList 
            Height          =   315
            ItemData        =   "frmMirage.frx":4BA44
            Left            =   120
            List            =   "frmMirage.frx":4BA46
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Action 
            Height          =   315
            ItemData        =   "frmMirage.frx":4BA48
            Left            =   2760
            List            =   "frmMirage.frx":4BA4A
            Style           =   2  'Dropdown List
            TabIndex        =   193
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command9 
            Caption         =   "x"
            Height          =   255
            Left            =   6480
            TabIndex        =   146
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton Btn 
            Caption         =   "S"
            Height          =   375
            Index           =   83
            Left            =   1680
            TabIndex        =   144
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "F"
            Height          =   375
            Index           =   70
            Left            =   2640
            TabIndex        =   143
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "G"
            Height          =   375
            Index           =   71
            Left            =   3120
            TabIndex        =   142
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "H"
            Height          =   375
            Index           =   72
            Left            =   3600
            TabIndex        =   141
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "J"
            Height          =   375
            Index           =   74
            Left            =   4080
            TabIndex        =   140
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "K"
            Height          =   375
            Index           =   75
            Left            =   4560
            TabIndex        =   139
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "L"
            Height          =   375
            Index           =   76
            Left            =   5040
            TabIndex        =   138
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "Z"
            Height          =   375
            Index           =   90
            Left            =   1440
            TabIndex        =   137
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "X"
            Height          =   375
            Index           =   88
            Left            =   1920
            TabIndex        =   136
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "C"
            Height          =   375
            Index           =   67
            Left            =   2400
            TabIndex        =   135
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "V"
            Height          =   375
            Index           =   86
            Left            =   2880
            TabIndex        =   134
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "B"
            Height          =   375
            Index           =   66
            Left            =   3360
            TabIndex        =   133
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "N"
            Height          =   375
            Index           =   78
            Left            =   3840
            TabIndex        =   132
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "M"
            Height          =   375
            Index           =   77
            Left            =   4320
            TabIndex        =   131
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "Q"
            Height          =   375
            Index           =   81
            Left            =   960
            TabIndex        =   130
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "W"
            Height          =   375
            Index           =   87
            Left            =   1440
            TabIndex        =   129
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "E"
            Height          =   375
            Index           =   69
            Left            =   1920
            TabIndex        =   128
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "R"
            Height          =   375
            Index           =   82
            Left            =   2400
            TabIndex        =   127
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "T"
            Height          =   375
            Index           =   84
            Left            =   2880
            TabIndex        =   126
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "Y"
            Height          =   375
            Index           =   89
            Left            =   3360
            TabIndex        =   125
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "U"
            Height          =   375
            Index           =   85
            Left            =   3840
            TabIndex        =   124
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "I"
            Height          =   375
            Index           =   73
            Left            =   4320
            TabIndex        =   123
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "O"
            Height          =   375
            Index           =   79
            Left            =   4800
            TabIndex        =   122
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "P"
            Height          =   375
            Index           =   80
            Left            =   5280
            TabIndex        =   121
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "D"
            Height          =   375
            Index           =   68
            Left            =   2160
            TabIndex        =   120
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Btn 
            Caption         =   "A"
            Height          =   375
            Index           =   65
            Left            =   1200
            TabIndex        =   119
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label35 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Action:"
            Height          =   255
            Left            =   2760
            TabIndex        =   192
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inventory item:"
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2.25000e5
      Left            =   14040
      Picture         =   "frmMirage.frx":4BA4C
      ScaleHeight     =   2.23636e5
      ScaleMode       =   0  'User
      ScaleWidth      =   477.091
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Mp3timer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   13320
      Top             =   8280
   End
   Begin VB.Timer NightTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13320
      Top             =   6840
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13320
      Top             =   7320
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13320
      Top             =   7800
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12840
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   7200
      Left            =   2130
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   13
      Top             =   390
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2160
      MaxLength       =   255
      TabIndex        =   11
      Top             =   7710
      Width           =   9480
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   13320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   990
      Left            =   2160
      TabIndex        =   0
      Top             =   8010
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1746
      _Version        =   393217
      BackColor       =   16512
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":1AB38E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBMP.VBMPlayer MusicPlayer 
      Height          =   1095
      Left            =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Height          =   1095
      Left            =   11760
      TabIndex        =   207
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   11760
      TabIndex        =   199
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   198
      Top             =   45
      Width           =   495
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   197
      Top             =   45
      Width           =   375
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2190
      TabIndex        =   196
      Top             =   45
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   144
      X2              =   776
      Y1              =   536
      Y2              =   536
   End
   Begin VB.Label Close 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Something?"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   74
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHat"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   73
      Top             =   960
      Width           =   2115
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AM / PM ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   72
      Top             =   960
      Width           =   825
   End
   Begin VB.Label lblcharstats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   0
      TabIndex        =   71
      Top             =   3075
      Width           =   2160
   End
   Begin VB.Label lblLabel20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   70
      Top             =   3120
      Width           =   2400
   End
   Begin VB.Label GameClock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   4560
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "It is now:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9435
      TabIndex        =   34
      Top             =   75
      Width           =   2250
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   -120
      TabIndex        =   26
      Top             =   6240
      Width           =   2280
   End
   Begin VB.Label lblWhosOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   11760
      TabIndex        =   16
      Top             =   960
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   11760
      TabIndex        =   12
      Top             =   4080
      Width           =   2040
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   0
      TabIndex        =   10
      Top             =   5160
      Width           =   2160
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1020
      Left            =   0
      TabIndex        =   9
      Top             =   915
      Width           =   2160
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB884B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5970
      TabIndex        =   8
      Top             =   75
      Width           =   2250
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2640
      TabIndex        =   7
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   2640
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00CB884B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   5970
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape shpTNL 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      Height          =   225
      Left            =   9435
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape bkHP 
      BackColor       =   &H00C6CEAD&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   2640
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C6CEAD&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   5970
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C6CEAD&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   9435
      Top             =   75
      Width           =   2250
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   2595
      Top             =   30
      Width           =   2325
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   5925
      Top             =   30
      Width           =   2340
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   9390
      Top             =   30
      Width           =   2340
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpellMemorized As Long

Private Sub back_Click()
Dim i As Long

    If currentsheet > 0 Then
        currentsheet = currentsheet - 1
        
        For i = 0 To 5
        Exp(i).Caption = vbNullString
        Level(i).Caption = vbNullString
        Picture1(i + 1).Visible = True
        
            If val(currentsheet * 5) + val(i + 1) <= MAX_SKILLS Then
                skillname(i).Caption = skill(val(currentsheet * 5) + val(i + 1)).Name
                    If val(skill(val(currentsheet * 5) + val(i + 1)).Pictop) = 0 And val(skill(val(currentsheet * 5) + val(i + 1)).Picleft) = 0 Then
                        Picture1(i + 1).Visible = False
                    Else
                        Exp(i).Caption = Player(MyIndex).SkilExp(val(currentsheet * 5) + val(i + 1))
                        Level(i).Caption = Player(MyIndex).SkilLvl(val(currentsheet * 5) + val(i + 1))
                        iconn(i).Left = -val(skill(val(currentsheet * 5) + val(i + 1)).Pictop * PIC_X)
                        iconn(i).Top = -val(skill(val(currentsheet * 5) + val(i + 1)).Picleft * PIC_Y)
                    End If
                skillname(i).Visible = True
            End If
        Next i
    End If
End Sub

Private Sub Btn_Click(Index As Integer)
Dim i As Long
Dim Text As String

    If ItemList.ListIndex > 0 Then
        If val(ReadINI("CONFIG", "Key" & Index & "_type", App.Path & "\config.ini")) <> 0 Then
            If MsgBox("There is already an action defined here, would you like to replace the existing one?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
        
        WriteINI "CONFIG", "Key" & Index & "_type", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Key" & Index & "_index", ItemList.ListIndex, App.Path & "\config.ini"
        WriteINI "CONFIG", "Key" & Index & "_button", Btn(Index).Caption, App.Path & "\config.ini"
        GoTo hell
    End If
    
    If Action.ListIndex > 0 Then
        If val(ReadINI("CONFIG", "Key" & Index & "_type", App.Path & "\config.ini")) <> 0 Then
            If MsgBox("There is already an action defined here do you want to replace the existing one?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
        
        WriteINI "CONFIG", "Key" & Index & "_type", 2, App.Path & "\config.ini"
        WriteINI "CONFIG", "Key" & Index & "_index", Action.ListIndex, App.Path & "\config.ini"
        WriteINI "CONFIG", "Key" & Index & "_button", Btn(Index).Caption, App.Path & "\config.ini"
        GoTo hell
    End If

Exit Sub
hell:
CS.ExecuteCmd cmCmdSelectAll
CS.SelText = vbNullString

For i = 65 To 90
    If val(ReadINI("CONFIG", "Key" & i & "_type", App.Path & "\config.ini")) = 1 Then
        Text = Text & ReadINI("CONFIG", "Key" & i & "_button", App.Path & "\config.ini") & " : " & ItemList.List(ReadINI("CONFIG", "Key" & i & "_index", App.Path & "\config.ini")) & vbCrLf
    End If

    If val(ReadINI("CONFIG", "Key" & i & "_type", App.Path & "\config.ini")) = 2 Then
        Text = Text & ReadINI("CONFIG", "Key" & i & "_button", App.Path & "\config.ini") & " : " & Action.List(ReadINI("CONFIG", "Key" & i & "_index", App.Path & "\config.ini")) & vbCrLf
    End If
Next i

CS.AddText (Text)
End Sub

Private Sub Close_Click()
    Call StopBGM
    Unload Me
End Sub

Private Sub AddDef_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.Value, App.Path & "\config.ini"
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.Value, App.Path & "\config.ini"
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, App.Path & "\config.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, App.Path & "\config.ini"
    If MyIndex <= 0 Then Exit Sub
    Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).Music))
End Sub

Private Sub cmdLeave_Click()
Dim packet As String
    packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdMember_Click()
Dim packet As String
    packet = "GUILDMEMBER" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub Command1_Click()
    picOptions.Visible = False
End Sub

Private Sub Command9_Click()
    BoxKeys.Visible = False
    CS.Visible = False
End Sub



Private Sub exit_Click()
    Skills.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long
Dim Ending As String
Dim Number As Long
Dim control As control

    For i = 1 To 5
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
       ' If FileExist("GUI\800X600" & Ending) Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\800X600" & Ending)
        If FileExist("GUI\Skill" & Ending) Then frmMirage.Skills.Picture = LoadPicture(App.Path & "\GUI\Skill" & Ending)
            If FileExist("GFX\Icons" & Ending) Then
                For j = 0 To 5
                    iconn(j).Picture = LoadPicture(App.Path & "\GFX\Icons" & Ending)
                Next j
            End If
    Next i
    
    Number = 1
    Do While Number < 0 + ReadINI("EMOS", "max", (App.Path & "\emo.ini"))
        If FileExist("gfx\emo" & Number & ".jpg") Then frmMirage.picEmote(Number - 1) = LoadPicture(App.Path & "\gfx\emo" & Number & ".jpg")
        Number = Number + 1
    Loop
    cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
    cClr.SetAsDefaultCursor Me.hWnd, True
    Dim lT As Long
      lT = Timer()
      ' do long operation
      Do While Timer - lT < 1
      Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub Label1_Click()
Dim i As Long

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If MouseDownX = GetPlayerX(i) And MouseDownY = GetPlayerY(i) Then
            Call SendData("playerchat" & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
    End If
Next i
End Sub

Private Sub Label13_Click()
' Set Their Guild Name and Their Rank
frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
Picture1(0).Visible = True
picInv3.Visible = False
picPlayerSpells.Visible = False
'picEquip.Visible = False
picWhosOnline.Visible = False
frmMirage.piccharstats.Visible = False
End Sub

Private Sub Label19_Click()
itmDesc.Visible = False
End Sub

Private Sub Label2_Click()
If picOptions.Visible = False Then
    picOptions.Visible = True
    chkplayername.Value = Trim$(ReadINI("CONFIG", "playername", App.Path & "\config.ini"))
chkplayerdamage.Value = Trim$(ReadINI("CONFIG", "playerdamage", App.Path & "\config.ini"))
chkplayerbar.Value = Trim$(ReadINI("CONFIG", "playerbar", App.Path & "\config.ini"))
chknpcname.Value = Trim$(ReadINI("CONFIG", "npcname", App.Path & "\config.ini"))
chknpcdamage.Value = Trim$(ReadINI("CONFIG", "npcdamage", App.Path & "\config.ini"))
chknpcbar.Value = Trim$(ReadINI("CONFIG", "npcbar", App.Path & "\config.ini"))
chkmusic.Value = Trim$(ReadINI("CONFIG", "music", App.Path & "\config.ini"))
chksound.Value = Trim$(ReadINI("CONFIG", "sound", App.Path & "\config.ini"))
chkbubblebar.Value = Trim$(ReadINI("CONFIG", "speechbubbles", App.Path & "\config.ini"))
chkAutoScroll.Value = Trim$(ReadINI("CONFIG", "AutoScroll", App.Path & "\config.ini"))
    picInv3.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    Picture1(0).Visible = False
    frmMirage.picGuildAdmin.Visible = False
    frmMirage.piccharstats.Visible = False
    Call UpdateVisInv
Else
    picOptions.Visible = False
    chkplayername.Value = Trim$(ReadINI("CONFIG", "playername", App.Path & "\config.ini"))
chkplayerdamage.Value = Trim$(ReadINI("CONFIG", "playerdamage", App.Path & "\config.ini"))
chkplayerbar.Value = Trim$(ReadINI("CONFIG", "playerbar", App.Path & "\config.ini"))
chknpcname.Value = Trim$(ReadINI("CONFIG", "npcname", App.Path & "\config.ini"))
chknpcdamage.Value = Trim$(ReadINI("CONFIG", "npcdamage", App.Path & "\config.ini"))
chknpcbar.Value = Trim$(ReadINI("CONFIG", "npcbar", App.Path & "\config.ini"))
chkmusic.Value = Trim$(ReadINI("CONFIG", "music", App.Path & "\config.ini"))
chksound.Value = Trim$(ReadINI("CONFIG", "sound", App.Path & "\config.ini"))
chkbubblebar.Value = Trim$(ReadINI("CONFIG", "speechbubbles", App.Path & "\config.ini"))
chkAutoScroll.Value = Trim$(ReadINI("CONFIG", "AutoScroll", App.Path & "\config.ini"))
    picInv3.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    Picture1(0).Visible = False
    frmMirage.picGuildAdmin.Visible = False
    frmMirage.piccharstats.Visible = False
    Call UpdateVisInv
End If
End Sub

Private Sub Label21_Click()
'picEquip.Visible = False
End Sub

Private Sub Label29_Click()
Dim i As Long

BoxKeys.Left = val(txtMyTextBox.Left + txtMyTextBox.Width) - val(BoxKeys.Width * 1.5) - 15
BoxKeys.Top = val(txtMyTextBox.Top - 152)

ItemList.addItem "None"
For i = 1 To MAX_INV
    If Player(MyIndex).Inv(i).num <> 0 Then
        ItemList.addItem item(Player(MyIndex).Inv(i).num).Name
    Else
        ItemList.addItem "Item" & i
    End If
Next i

ItemList.ListIndex = 0
Action.ListIndex = 0

BoxKeys.Visible = True
CS.Visible = True
End Sub

Private Sub Label3_Click()
Call GameDestroy
End Sub


Private Sub Label30_Click()
Dim i As Long
currentsheet = 0

For i = 0 To 5
Exp(i).Caption = vbNullString
Level(i).Caption = vbNullString
Picture1(i + 1).Visible = True
    If val(currentsheet * 5) + val(i + 1) <= MAX_SKILLS Then
        skillname(i).Caption = skill(val(currentsheet * 5) + val(i + 1)).Name
            If val(skill(val(currentsheet * 5) + val(i + 1)).Pictop) = 0 And val(skill(val(currentsheet * 5) + val(i + 1)).Picleft) = 0 Then
                Picture1(i + 1).Visible = False
            Else
                Exp(i).Visible = True
                Level(i).Visible = True
                Exp(i).Caption = Player(MyIndex).SkilExp(val(currentsheet * 5) + val(i + 1))
                Level(i).Caption = Player(MyIndex).SkilLvl(val(currentsheet * 5) + val(i + 1))
                iconn(i).Left = -val(skill(val(currentsheet * 5) + val(i + 1)).Pictop * PIC_X)
                iconn(i).Top = -val(skill(val(currentsheet * 5) + val(i + 1)).Picleft * PIC_Y)
            End If
        skillname(i).Visible = True
    End If
Next i

Skills.Visible = True
End Sub

Private Sub Label39_Click()
picWhosOnline.Visible = False
picInv3.Visible = False
'picEquip.Visible = False
Picture1(0).Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = False
picOptions.Visible = False
Dim i As Long

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If MouseDownX = GetPlayerX(i) And MouseDownY = GetPlayerY(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            Call SendData("playerchat" & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
    End If
Next i
End Sub

Private Sub Label40_Click()
frmMirage.Visible = False
Call TcpDestroy
Call TcpInit
frmMainMenu.Show
End Sub

Private Sub Label41_Click()
picInv3.Visible = False
itmDesc.Visible = False
picrciinfo.Visible = False
End Sub

Private Sub Label7_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
If picInv3.Visible = False Then
    Call UpdateVisInv
    picInv3.Visible = True
    Picture1(0).Visible = False
 '   picEquip.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    ''picEquip.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
Else
    picInv3.Visible = False
    Picture1(0).Visible = False
  '  picEquip.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    ''picEquip.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
End If
End Sub



Private Sub Label8_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
If frmMirage.picPlayerSpells.Visible = False Then
    Call SendData("spells" & SEP_CHAR & END_CHAR)
    picInv3.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    'picEquip.Visible = False
    Picture1(0).Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
Else
    picPlayerSpells.Visible = False
    picInv3.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    'picEquip.Visible = False
    Picture1(0).Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
End If
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
End Sub

Private Sub lblcharstats_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
If frmMirage.piccharstats.Visible = False Then
picWhosOnline.Visible = False
picInv3.Visible = False
Picture1(0).Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = True
Else
picWhosOnline.Visible = False
picInv3.Visible = False
Picture1(0).Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = False
End If
End Sub




Private Sub lblForgetSpell_Click()
If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
If MsgBox("Are you sure you want to forget this spell?", vbYesNo, "Forget Spell") = vbNo Then Exit Sub
Call SendData("forgetspell" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
picPlayerSpells.Visible = False
End If
Else
Call AddText("No spell here.", BrightRed)
End If
End Sub

Private Sub lblLabel20_Click()
    InGame = False
End Sub

Private Sub lblrcidrop_Click()
    Call DropItems
End Sub

Private Sub lblrcidrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub lblrcigift_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
Call GiveItems
End Sub

Private Sub lblrcigift_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub lblrciinfo_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
Call show_info(RCIINDEX)
End Sub

Private Sub lblrciinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub lblrciuse_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = True
Call useitem(RCIINDEX)
End Sub

Private Sub lblrciuse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub lblSTATWINDOW_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub lblWhosOnline_Click()
If picWhosOnline.Visible = False Then
Call SendOnlineList
picWhosOnline.Visible = True
picInv3.Visible = False
'picEquip.Visible = False
Picture1(0).Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = False
picOptions.Visible = False
Else
picWhosOnline.Visible = False
picInv3.Visible = False
'picEquip.Visible = False
Picture1(0).Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = False
picOptions.Visible = False
End If

End Sub

Private Sub lbrcchat_Click()
Call SendData("playerchat" & SEP_CHAR & GetPlayerName(RWINDEX) & SEP_CHAR & END_CHAR)
frmMirage.picrclick.Visible = False
End Sub

Private Sub lbrcparty_Click()
Call SendPartyRequest(GetPlayerName(RWINDEX))
frmMirage.picrclick.Visible = False
End Sub

Private Sub lbrcpm_Click()
frmMirage.txtMyTextBox.Text = "!" & GetPlayerName(RWINDEX) & " "
frmMirage.picrclick.Visible = False

frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub lbrctrade_Click()
Call SendTradeRequest(GetPlayerName(RWINDEX))
frmMirage.picrclick.Visible = False
End Sub


Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim$(lstOnline.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstSpells_DblClick()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        SpellMemorized = lstSpells.ListIndex + 1
        Call AddText("Successfully memorized spell!", White)
    Else
        Call AddText("No spell here to memorize.", BrightRed)
    End If
End Sub




Private Sub MENUreloader_Timer()

frmCustom1.Visible = True

End Sub


Private Sub Mousetmr_Timer()
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
   Mousetmr.Enabled = False
End Sub

Private Sub next_Click()
Dim i As Long

    If currentsheet <= Int(val(MAX_SKILLS / 6) - 1) Then
        currentsheet = currentsheet + 1
        
        For i = 0 To 5
        Exp(i).Caption = vbNullString
        Level(i).Caption = vbNullString
        Picture1(i + 1).Visible = True
        
            If val(currentsheet * 6) + val(i + 1) <= MAX_SKILLS Then
                skillname(i).Caption = skill(val(currentsheet * 6) + val(i + 1)).Name
                    If val(skill(val(currentsheet * 6) + val(i + 1)).Pictop * PIC_X) = 0 And val(skill(val(currentsheet * 6) + val(i + 1)).Picleft * PIC_Y) = 0 Then
                        Picture1(i + 1).Visible = False
                    Else
                        Exp(i).Caption = Player(MyIndex).SkilExp(val(currentsheet * 6) + val(i + 1))
                        Level(i).Caption = Player(MyIndex).SkilLvl(val(currentsheet * 6) + val(i + 1))
                        iconn(i).Left = -val(skill(val(currentsheet * 6) + val(i + 1)).Pictop * PIC_X)
                        iconn(i).Top = -val(skill(val(currentsheet * 6) + val(i + 1)).Picleft * PIC_Y)
                    End If
                skillname(i).Visible = True
            Else
                Picture1(i + 1).Visible = False
            End If
        Next i
    End If
End Sub

Private Sub picInv_DblClick(Index As Integer)
Dim d As Long
frmMirage.picrciinfo.Visible = False
If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
        If item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            picInv(d - 1).Picture = LoadPicture()
        End If
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub useitem(Index As Integer)
Dim d As Long
If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
        If item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            picInv(d - 1).Picture = LoadPicture()
        End If
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim AMT As Integer
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor picInv(Index).hWnd, True
   Mousetmr.Enabled = True
    Inventory = Index + 1
    frmMirage.SelectedItem.Visible = True
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
    If frmNewShop.fixItems And frmNewShop.Visible = True Then
        'Fix the selected item
        frmNewShop.FixItem (GetPlayerInvItemNum(MyIndex, Inventory))
    Else
        'We're selling items to a shop
        If frmNewShop.SellItems And frmNewShop.Visible = True Then
            If item(GetPlayerInvItemNum(MyIndex, Inventory)).Stackable = YES Then
                AMT = val(InputBox("How many would you like to sell?", "Sell Items")) + 0
                If AMT > 0 Then
                    'Sell the items
                    frmNewShop.Buyback GetPlayerInvItemNum(MyIndex, Inventory), Inventory, AMT
                Else
                    'ADD ANY TEXT YOU WANT HERE. BLANK BY DEFAULT.
                    'Call MsgBox(blah)
                End If
            Else
                'Sell the selected item
                frmNewShop.Buyback GetPlayerInvItemNum(MyIndex, Inventory), Inventory
            End If
        Else
            'Regular click
            If Button = 1 Then
                Call UpdateVisInv
            ElseIf Button = 2 Then
                Call DropItems
            End If
        End If
    End If
    frmMirage.picrciinfo.Visible = False
    ElseIf Button = 2 Then
        If Player(MyIndex).Inv(Index + 1).num <> 0 Then
        RCIINDEX = Index
        frmMirage.lblrciname.Caption = Trim$(item(Player(MyIndex).Inv(Index + 1).num).Name)
        frmMirage.picrciinfo.Top = Y + frmMirage.picInv3.Top + frmMirage.picInv(Index).Top
        frmMirage.picrciinfo.Left = X + frmMirage.picInv3.Left + frmMirage.picInv(Index).Left
        frmMirage.picrciinfo.Visible = True
        End If
    End If
End Sub

Private Sub show_info(Index As Integer)
Dim d As Long
d = Index

cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
    If Player(MyIndex).Inv(d + 1).num > 0 Then
        If item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            If Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
                itmDesc.Height = 17
                itmDesc.Top = 150
            Else
                itmDesc.Height = 289
                itmDesc.Top = 35
            End If
        Else
            If Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
                itmDesc.Height = 161
                itmDesc.Top = 96
            Else
                itmDesc.Height = 289
                itmDesc.Top = 35
            End If
        End If
        If item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            Else
                descName.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
            End If
        End If
        
        descStr.Caption = item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " " & STAT1
        descDef.Caption = item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " " & STAT2
        descSpeed.Caption = item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " " & STAT4
        descHpMp.Caption = "HP: " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = STAT1 & ": " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " " & STAT2 & ": " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = STAT3 & ": " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " " & STAT4 & ": " & item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        desc.Caption = Trim$(item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long

    Call SendData("key" & SEP_CHAR & KeyCode & SEP_CHAR & END_CHAR)

    Call CheckInput(0, KeyCode, Shift)
    If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmadmin.Visible = False
            frmadmin.Visible = True
        End If
    End If
    
If KeyCode = vbKeyF2 Then
For i = 1 To MAX_INV

    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDHP Then
            Call SendUseItem(i)
            Call AddText("You restored you Health", Yellow)
            Exit Sub
        End If
    Else
            If i = MAX_INV Then Call AddText("You dont have anything to restored your Health!", Red)
    End If
    Next i
End If

If KeyCode = vbKeyF3 Then
For i = 1 To MAX_INV

    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDMP Then
        Call SendUseItem(i)
        Call AddText("You restored you Mana", Yellow)
        
        Exit Sub
        End If
    Else
        If i = MAX_INV Then Call AddText("You dont have anything to restored your Mana!", Red)
    End If
    Next i
End If
    
    ' The Guild Creator
If KeyCode = vbKeyF4 Then
    If Player(MyIndex).Access > 0 Then
        'frmGuild.Show vbModeless, frmMirage
    End If
End If

If KeyCode = vbKeyPageUp Then
    Call SendHotScript1
End If
    
If KeyCode = vbKeyDelete Then
    Call SendHotScript2
End If
    
If KeyCode = vbKeyEnd Then
    Call SendHotScript3
End If
    
If KeyCode = vbKeyPageDown Then
    Call SendHotScript4
End If

' The Guild Maker
If KeyCode = vbKeyF5 Then
    frmMirage.picGuildAdmin.Visible = True
    frmMirage.picInv3.Visible = False
    frmMirage.Picture1(0).Visible = False
    'frmMirage.picEquip.Visible = False
    frmMirage.picPlayerSpells.Visible = False
    frmMirage.picWhosOnline.Visible = False
End If
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & SpellMemorized & SEP_CHAR & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Cannot cast while walking!", BrightRed)
                End If
            End If
        Else
            Call AddText("No spell here memorized.", BrightRed)
        End If
    Else
        Call CheckInput(0, KeyCode, Shift)
    End If
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        i = 0
        ii = 0
        Do
            If FileExist("Screenshot" & i & ".bmp") = True Then
                i = i + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp")
                ii = 1
            End If
            
            DoEvents
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        If Screen_RESIZED = 1 Then
            ScreenShot.Picture = CaptureArea(frmMirage, picUber.Left, picUber.Top, picUber.Width - 8, picUber.Height)
        Else
            ScreenShot.Picture = CaptureArea(frmMirage, picUber.Left, picUber.Top, picUber.Width, picUber.Height)
        End If
        i = 0
        ii = 0
        Do
            If FileExist("Screenshot" & i & ".bmp") = True Then
                i = i + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp")
                ii = 1
            End If
            
            DoEvents
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyHome Then
    d = GetPlayerDir(MyIndex)
        If Player(MyIndex).Moving = NO Then
                If Player(MyIndex).Dir = DIR_DOWN Then
                    Call SetPlayerDir(MyIndex, DIR_LEFT)
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                    Call SetPlayerDir(MyIndex, DIR_UP)
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                ElseIf Player(MyIndex).Dir = DIR_UP Then
                    Call SetPlayerDir(MyIndex, DIR_RIGHT)
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                    Call SetPlayerDir(MyIndex, DIR_DOWN)
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                End If
        End If
    End If

    If KeyCode < 91 And KeyCode > 64 Then
        If val(ReadINI("CONFIG", "Key" & KeyCode & "_type", App.Path & "\config.ini")) = 1 Then
            Call SendUseItem(ReadINI("CONFIG", "Key" & KeyCode & "_type", App.Path & "\config.ini"))
        End If
        
        If val(ReadINI("CONFIG", "Key" & KeyCode & "_type", App.Path & "\config.ini")) = 2 Then
            Call SendScript(ReadINI("CONFIG", "Key" & KeyCode & "_type", App.Path & "\config.ini"))
        End If
    End If

End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor picInv(Index).hWnd, True
End Sub

Private Sub picInv3_Click()
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor picInv3.hWnd, True
   Mousetmr.Enabled = True
frmMirage.picrciinfo.Visible = False
End Sub

Private Sub picInv3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
cClr.SetAsDefaultCursor picInv3.hWnd, True
End Sub

Private Sub PicOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, X, Y)
End Sub

Private Sub picrciinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor Me.hWnd, True
End Sub

Private Sub picrclick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor picrclick.hWnd, True
   Mousetmr.Enabled = True
End Sub

Private Sub picrclick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor picrclick.hWnd, True
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub


Private Sub picUber_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub


Private Sub picUber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim xcalc As Single
Dim ycalc As Single

cClr.LoadFromFile App.Path & "\GFX\mousedown.cur", , True
   cClr.SetAsDefaultCursor picUber.hWnd, True
   Mousetmr.Enabled = True
frmMirage.picrciinfo.Visible = False
X = X / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
Y = Y / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

xcalc = (NewPlayerX * PIC_X) / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
ycalc = (NewPlayerY * PIC_Y) / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

If 0 + ReadINI("CONFIG", "mousebug", App.Path & "\config.ini") = 1 Then
    Call WriteINI("CONFIG", "Res", 1, (App.Path & "\config.ini"))
    xcalc = xcalc + 1
    ycalc = ycalc + 1
End If

    If Button = 2 Then

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If ((X + xcalc) / PIC_X) = GetPlayerX(i) And ((Y + ycalc) / PIC_Y) = GetPlayerY(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            picrclick.Visible = True
            picrclick.Top = X
            picrclick.Left = Y
            Exit Sub
        End If
    End If
Next i
    End If
    
    If (Button = 1 Or Button = 2) And InEditor = True Then
        'Call AddText("Clicked xcalc" & (x + xcalc) & " ycalc" & (y + ycalc), BrightRed)
        Call EditorMouseDown(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
    End If
    
    If (Button = 1 Or Button = 2) And InEditor = False Then
        'Call AddText("Clicked xcalc" & (x + xcalc) & " ycalc" & (y + ycalc), BrightRed)
        Call PlayerSearch(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
        'Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), ycalc)
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        'Call AddText("Clicked xcalc" & (x + xcalc) & " ycalc" & (y + ycalc), BrightRed)
        Call HouseEditorMouseDown(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
        'Call HouseEditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), ycalc)
    End If
    
    If (Button = 1 Or Button = 2) And CanonUsed = 1 Then
        Call CanonShoot(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
    End If
    
End Sub

Private Sub picUber_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xcalc As Single
Dim ycalc As Single

X = X / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
Y = Y / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

xcalc = (NewPlayerX * PIC_X) / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
ycalc = (NewPlayerY * PIC_Y) / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

If 0 + ReadINI("CONFIG", "mousebug", App.Path & "\config.ini") = 1 Then
    Call WriteINI("CONFIG", "Res", 1, (App.Path & "\config.ini"))
    xcalc = xcalc + 1
    ycalc = ycalc + 1
End If

cClr.LoadFromFile App.Path & "\GFX\mousemove.cur", , True
   cClr.SetAsDefaultCursor picUber.hWnd, True


    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        Call HouseEditorMouseDown(Button, Shift, Int((X + (xcalc))), Int((Y + (ycalc))))
    End If

    CanonX = Int(X)
    CanonY = Int(Y)
    
    frmMapEditor.Caption = "Map Editor - " & "X: " & Int(Int((X + (xcalc))) / 32) & " Y: " & Int(Int((Y + (ycalc))) / 32)
    frmHouseEditor.Caption = "House Editor - " & "X: " & Int(Int((X + (xcalc))) / 32) & " Y: " & Int(Int((Y + (ycalc))) / 32)

' (x + (NewPlayerX * PIC_X)) BECAME Int((x + (xcalc)))
' Int((y + (ycalc)))

End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        Call HouseEditorMouseDown(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If

    CanonX = Int(X / (frmMirage.picUber.Width / frmMirage.picScreen.Width))
    CanonY = Int(Y / (frmMirage.picUber.Height / frmMirage.picScreen.Height))


    frmMapEditor.Caption = "Map Editor - " & "X: " & Int((X + (NewPlayerX * PIC_X)) / 32) & " Y: " & Int((Y + (NewPlayerY * PIC_Y)) / 32)
    frmHouseEditor.Caption = "House Editor - " & "X: " & Int((X + (NewPlayerX * PIC_X)) / 32) & " Y: " & Int((Y + (NewPlayerY * PIC_Y)) / 32)

End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    If (Button = 1 Or Button = 2) And InEditor = True Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
        Call EditorMouseDown(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If
    
    If Button = 1 And InEditor = False Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
        Call PlayerSearch(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If
    
        If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
 
        Call HouseEditorMouseDown(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If
End Sub



Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    MAX_BLT_LINE = scrlBltText.Value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "On Screen Text Line Amount: " & scrlBltText.Value
End Sub


Private Sub ScrlResolution_Change()
If 0 + ReadINI("CONFIG", "mousebug", App.Path & "\config.ini") = 1 And InEditor = True Then Exit Sub
Call WriteINI("CONFIG", "Res", frmMirage.ScrlResolution.Value, (App.Path & "\config.ini"))
Screen_RESIZED = 0
End Sub



Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub


Private Sub tmrGameClock_Timer()
    IncrementGameClock
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then
        tmrRainDrop.Enabled = False
        Exit Sub
    End If
    If BLT_RAIN_DROPS > 0 Then
        If DropRain(BLT_RAIN_DROPS).Randomized = False Then
            Call RNDRainDrop(BLT_RAIN_DROPS)
        End If
    End If
    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1
    If tmrRainDrop.Interval > 30 Then
        tmrRainDrop.Interval = tmrRainDrop.Interval - 10
    End If
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then
        tmrSnowDrop.Enabled = False
        Exit Sub
    End If
    If BLT_SNOW_DROPS > 0 Then
        If DropSnow(BLT_SNOW_DROPS).Randomized = False Then
            Call RNDSnowDrop(BLT_SNOW_DROPS)
        End If
    End If
    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1
    If tmrSnowDrop.Interval > 30 Then
        tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
    End If
End Sub

Private Sub txtChat_GotFocus()
    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub picInv3entory_Click()
    picInv3.Visible = True
End Sub

Private Sub lblUseItem_Click()
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
        If item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            picInv(d - 1).Picture = LoadPicture()
        End If
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub lblDropItem_Click()
    Call DropItems
End Sub

Sub DropItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done
If Inventory <= 0 Then Exit Sub

    InvNum = Inventory
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).Bound = 0 Then
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
            GoldAmount = InputBox("How much " & Trim$(item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim$(item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then
                Call SendDropItem(InvNum, GoldAmount)
            End If
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
        MsgBox "The variable cant handle that amount!"
    End If
End Sub


Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & SpellMemorized & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub lblCancel_Click()
    picInv3.Visible = False
End Sub

Private Sub lblSpellsCancel_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub picSpells_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picStats_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picTrade_Click()
    Call SendData("trade" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picQuit_Click()
    InGame = False
End Sub

Private Sub cmdAccess_Click()
Dim packet As String

    packet = "GUILDCHANGEACCESS" & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdDisown_Click()
Dim packet As String

    packet = "GUILDDISOWN" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdTrainee_Click()
Dim packet As String
    
    packet = "GUILDTRAINEE" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub picOffline_Click()
    Call SendOnlineList
    lstOnline.Visible = False
    'Label9.Visible = False
End Sub

Private Sub picOnline_Click()
    Call SendOnlineList
    lstOnline.Visible = True
    'Label9.Visible = True
End Sub



Private Sub lstSpells_GotFocus()
'picScreen.SetFocus
picUber.SetFocus
End Sub



