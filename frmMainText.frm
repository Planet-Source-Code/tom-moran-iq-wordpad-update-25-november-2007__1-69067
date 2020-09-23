VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainText 
   AutoRedraw      =   -1  'True
   Caption         =   "iQ WordPad"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9945
   ClipControls    =   0   'False
   Icon            =   "frmMainText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   7680
      ScaleHeight     =   1245
      ScaleWidth      =   1425
      TabIndex        =   30
      Top             =   2940
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox RulerBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F3CDB4&
      Height          =   475
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   15300
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   840
      Width           =   15360
      Begin VB.PictureBox picRuler 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E8B08A&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   120
         Picture         =   "frmMainText.frx":0E42
         ScaleHeight     =   435
         ScaleWidth      =   18000
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   60
         Width           =   18000
         Begin VB.Image imgRight 
            Height          =   210
            Left            =   10200
            Picture         =   "frmMainText.frx":1049C
            Top             =   150
            Width           =   140
         End
         Begin VB.Image imgMargin 
            Height          =   225
            Left            =   10260
            Picture         =   "frmMainText.frx":109C6
            Top             =   0
            Width           =   14310
         End
         Begin VB.Image imgTab 
            Height          =   135
            Index           =   0
            Left            =   1440
            Picture         =   "frmMainText.frx":1B1D8
            Top             =   100
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Image imgLeft 
            Height          =   210
            Left            =   120
            Picture         =   "frmMainText.frx":1B6B2
            Top             =   150
            Width           =   135
         End
         Begin VB.Image imgHang 
            Height          =   120
            Left            =   120
            Picture         =   "frmMainText.frx":1BBEC
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin VB.PictureBox picInsert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   8940
      ScaleHeight     =   525
      ScaleWidth      =   585
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picFormatbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "frmMainText.frx":1C0C6
      ScaleHeight     =   450
      ScaleWidth      =   15360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   15360
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Combo2"
         ToolTipText     =   " Font "
         Top             =   60
         Width           =   2055
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Combo1"
         ToolTipText     =   " Font Size "
         Top             =   60
         Width           =   795
      End
      Begin iQWordPad.CandyButton butColor 
         Height          =   345
         Left            =   4680
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   " Font Color "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":32908
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton butHighlight 
         Height          =   345
         Left            =   8865
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   " Highlight Text "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":32EA2
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton butSpacing 
         Height          =   345
         Left            =   8370
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   " Line Spacing "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":3343C
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton butCaseChange 
         Height          =   345
         Left            =   9390
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   " Change Case  "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":339E6
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton butRecentFonts 
         Height          =   345
         Left            =   7875
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   " Recent Fonts List "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":33F90
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin VB.Image butnumber 
         Height          =   330
         Index           =   3
         Left            =   7340
         Picture         =   "frmMainText.frx":3453A
         ToolTipText     =   " Bullet Styles  "
         Top             =   50
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Image butnumber 
         Height          =   330
         Index           =   2
         Left            =   7340
         Picture         =   "frmMainText.frx":34894
         ToolTipText     =   " Bullet Styles  "
         Top             =   50
         Width           =   165
      End
      Begin VB.Image butnumber 
         Height          =   330
         Index           =   1
         Left            =   7020
         Picture         =   "frmMainText.frx":34BEE
         ToolTipText     =   " Numbered "
         Top             =   50
         Width           =   345
      End
      Begin VB.Image butnumber 
         Height          =   330
         Index           =   0
         Left            =   7020
         Picture         =   "frmMainText.frx":35260
         ToolTipText     =   " Numbered "
         Top             =   50
         Width           =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   7740
         X2              =   7740
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   7725
         X2              =   7725
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Image butbullet 
         Height          =   330
         Index           =   1
         Left            =   6600
         Picture         =   "frmMainText.frx":358D2
         ToolTipText     =   " Bullet "
         Top             =   50
         Width           =   345
      End
      Begin VB.Image butbullet 
         Height          =   330
         Index           =   0
         Left            =   6600
         Picture         =   "frmMainText.frx":35F44
         ToolTipText     =   " Bullet "
         Top             =   50
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   5
         Left            =   6060
         Picture         =   "frmMainText.frx":365B6
         ToolTipText     =   " Align Right "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   4
         Left            =   6060
         Picture         =   "frmMainText.frx":36C28
         ToolTipText     =   " Align Right "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   3
         Left            =   5640
         Picture         =   "frmMainText.frx":3729A
         ToolTipText     =   " Align Center "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   2
         Left            =   5640
         Picture         =   "frmMainText.frx":3790C
         ToolTipText     =   " Align Center "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   1
         Left            =   5220
         Picture         =   "frmMainText.frx":37F7E
         ToolTipText     =   " Align Left "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butAlign 
         Height          =   330
         Index           =   0
         Left            =   5220
         Picture         =   "frmMainText.frx":385F0
         ToolTipText     =   " Align Left "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butUnderline 
         Height          =   330
         Index           =   1
         Left            =   4200
         Picture         =   "frmMainText.frx":38C62
         ToolTipText     =   " Underline "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butUnderline 
         Height          =   330
         Index           =   0
         Left            =   4200
         Picture         =   "frmMainText.frx":392D4
         ToolTipText     =   " Underline "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butItalic 
         Height          =   330
         Index           =   1
         Left            =   3780
         Picture         =   "frmMainText.frx":39946
         ToolTipText     =   " Italic "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butItalic 
         Height          =   330
         Index           =   0
         Left            =   3780
         Picture         =   "frmMainText.frx":39FB8
         ToolTipText     =   " Italic "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butBold 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3360
         Picture         =   "frmMainText.frx":3A62A
         ToolTipText     =   " Bold "
         Top             =   60
         Width           =   345
      End
      Begin VB.Image butBold 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   3360
         Picture         =   "frmMainText.frx":3AC9C
         ToolTipText     =   " Bold "
         Top             =   60
         Width           =   345
      End
   End
   Begin VB.Timer tmrClipboard 
      Interval        =   50
      Left            =   9420
      Top             =   2820
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "frmMainText.frx":3B30E
      ScaleHeight     =   390
      ScaleWidth      =   9945
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5520
      Width           =   9945
      Begin VB.PictureBox picLnCol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6240
         Picture         =   "frmMainText.frx":53950
         ScaleHeight     =   390
         ScaleWidth      =   4485
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   4485
         Begin VB.Image imgGripper 
            Height          =   225
            Left            =   4220
            MousePointer    =   8  'Size NW SE
            Picture         =   "frmMainText.frx":594FA
            Top             =   150
            Width           =   255
         End
         Begin VB.Label lblInsNumCap 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NUM"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1485
            TabIndex        =   22
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblInsNumCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CAP"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   21
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblInsNumCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INSERT"
            Height          =   195
            Index           =   0
            Left            =   140
            TabIndex        =   20
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Ln 1, Col 1"
            Height          =   195
            Left            =   2100
            TabIndex        =   16
            Top             =   120
            Width           =   1995
         End
      End
      Begin VB.Label lblIQFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Untitled"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   5700
      End
   End
   Begin RichTextLib.RichTextBox txtInsert 
      Height          =   1035
      Left            =   9720
      TabIndex        =   13
      Top             =   1740
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1826
      _Version        =   393217
      RightMargin     =   2.00000e5
      TextRTF         =   $"frmMainText.frx":59848
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "frmMainText.frx":598C7
      ScaleHeight     =   450
      ScaleWidth      =   9945
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9945
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   " New "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":70109
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   1
         Left            =   660
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   " Open... "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":70B1B
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   " Save "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":710B5
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   3
         Left            =   1860
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   " Print "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":7164F
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   4
         Left            =   3240
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   " Cut "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":719E9
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   5
         Left            =   3840
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   " Copy "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":71F83
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   6
         Left            =   4440
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " Paste "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":7251D
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   7
         Left            =   5040
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   " Delete "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":72F2F
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   8
         Left            =   6240
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   " Undo/Redo  "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":734C9
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   9
         Left            =   5640
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " Find/Replace "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":73A63
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   10
         Left            =   7020
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " Insert Picture "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":73FFD
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   11
         Left            =   7620
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " Insert Date/Time  "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":745A7
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   12
         Left            =   2460
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   " Print Preview "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":74B41
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   13
         Left            =   8700
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   " Spell Check "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":750EB
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   14
         Left            =   8160
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   " Insert Symbol "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":75685
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQWordPad.CandyButton butReferenes 
         Height          =   345
         Left            =   9240
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   " References "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "frmMainText.frx":75C1F
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000003&
         X1              =   6825
         X2              =   6825
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   6840
         X2              =   6840
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000003&
         X1              =   3060
         X2              =   3060
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   3075
         X2              =   3075
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         Visible         =   0   'False
         X1              =   9915
         X2              =   9915
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         Visible         =   0   'False
         X1              =   9900
         X2              =   9900
         Y1              =   15
         Y2              =   430
      End
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4635
      Left            =   300
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8176
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmMainText.frx":761C9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   8160
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   $"frmMainText.frx":76242
   End
   Begin VB.Image imgRuler 
      Height          =   315
      Index           =   1
      Left            =   7800
      Picture         =   "frmMainText.frx":762D0
      Top             =   4860
      Visible         =   0   'False
      Width           =   15030
   End
   Begin VB.Image imgRuler 
      Height          =   315
      Index           =   0
      Left            =   7800
      Picture         =   "frmMainText.frx":859D2
      Top             =   4500
      Visible         =   0   'False
      Width           =   15000
   End
   Begin VB.Menu xxxmnuFile 
      Caption         =   "File"
      Begin VB.Menu xxxmnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu xxxmnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu xxxmnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu xxxmnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu xxSepF1 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuPageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu xxxmnuPrinterSetup 
         Caption         =   "Printer Setup..."
      End
      Begin VB.Menu xxxmnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu xxxmnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu xxmnuSepF2 
         Caption         =   "-"
      End
      Begin VB.Menu xxmnuRFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuRecentFiles 
            Caption         =   " Do Not Show"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles1"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles8"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu xxmnuSepF5 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSend 
         Caption         =   "Send..."
      End
      Begin VB.Menu xxxmnuSepF7 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu xxxmnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu xxxmnuUndo 
         Caption         =   "Undo                      Ctrl+Z"
      End
      Begin VB.Menu xxmnuSepE4 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuCut 
         Caption         =   "Cut to clipboard     Ctrl+X"
      End
      Begin VB.Menu xxxmnuCopy 
         Caption         =   "Copy                      Ctrl+C"
      End
      Begin VB.Menu xxxmnuPaste 
         Caption         =   "Paste                     Ctrl+V"
      End
      Begin VB.Menu xxxmnuDelete 
         Caption         =   "Delete                    Del"
      End
      Begin VB.Menu xxmnuSepE2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuFind 
         Caption         =   "Find...                    Ctrl F"
      End
      Begin VB.Menu xxxmnuReplace 
         Caption         =   "Replace...              Ctrl+H"
      End
      Begin VB.Menu xxxmnuGoto 
         Caption         =   "Go To....                Ctrl+G"
      End
      Begin VB.Menu xxmnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSelectAll 
         Caption         =   "Select All                Ctrl+A"
      End
   End
   Begin VB.Menu xxxmnuView 
      Caption         =   "View"
      Begin VB.Menu xxxmnuTool 
         Caption         =   "Tool Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu xxxmnuFormatBar 
         Caption         =   "Format Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu xxxmnuStatus 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu xxxmnuRuler 
         Caption         =   "Ruler"
         Checked         =   -1  'True
      End
      Begin VB.Menu xxmnuSepV1 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSpellCheck 
         Caption         =   "Spell Check"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuRefer 
         Caption         =   "References"
         Begin VB.Menu mnuReferences 
            Caption         =   "Dictionary (Merriam-Webster)"
            Index           =   0
         End
         Begin VB.Menu mnuReferences 
            Caption         =   "Thesaurus  (Roget's)"
            Index           =   1
         End
         Begin VB.Menu mnuReferences 
            Caption         =   "Encyclopedia (Wikipedia)"
            Index           =   2
         End
         Begin VB.Menu mnuReferences 
            Caption         =   "Internet Search (Google)"
            Index           =   3
         End
      End
      Begin VB.Menu xxxmnuWordCount 
         Caption         =   "Document Statistics"
         Shortcut        =   ^D
      End
      Begin VB.Menu xxxmnuTextProperties 
         Caption         =   "File Properties"
      End
      Begin VB.Menu xxmnuSepV2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuRulerStyle 
         Caption         =   "Ruler Measurement"
         Begin VB.Menu xxxmnuRulerOption 
            Caption         =   "Inches"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu xxxmnuRulerOption 
            Caption         =   "Centimeters"
            Index           =   1
         End
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
         Shortcut        =   ^K
      End
      Begin VB.Menu xxxmnuViewClipboard 
         Caption         =   "Clipboard"
      End
   End
   Begin VB.Menu xxxmnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu xxxmnuInsertPic 
         Caption         =   "Insert Picture..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu xxxmnuInsDocument 
         Caption         =   "Insert Document..."
      End
      Begin VB.Menu xxxmnuSepI1 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuTable 
         Caption         =   "Table..."
      End
      Begin VB.Menu xxxmnuSepI6 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuCharMap 
         Caption         =   "Symbols/Character..."
      End
      Begin VB.Menu xxxmnuTimeDate 
         Caption         =   "Date and Time..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu xxxmnuFormat 
      Caption         =   "Format"
      Begin VB.Menu xxxmnuFont 
         Caption         =   "Font..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu xxxmnuRecentFonts 
         Caption         =   "Recent Fonts"
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "    ---  Recent Fonts List  ---           "
            Index           =   0
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "Arial"
            Index           =   1
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFonts 
            Caption         =   "RecentFonts8"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu xxxmnuSetAttrib 
         Caption         =   "Attributes"
         Begin VB.Menu mnuAttributes 
            Caption         =   "Text Color..."
            Index           =   0
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "Bold"
            Index           =   2
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "Italic"
            Index           =   3
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "Underline"
            Index           =   4
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "Strikethru"
            Index           =   5
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuAttributes 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu xxxmnuSuperScript 
            Caption         =   "SuperScript"
            Shortcut        =   +{F7}
         End
         Begin VB.Menu xxxmnuSubScript 
            Caption         =   "SubScript"
            Shortcut        =   +{F8}
         End
      End
      Begin VB.Menu mnuChangeCase 
         Caption         =   "Change Case"
         Begin VB.Menu mnuCaseChange 
            Caption         =   "Sentence case"
            Index           =   0
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "lower case"
            Index           =   1
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "UPPER CASE"
            Index           =   2
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "Capitalize Each Word"
            Index           =   3
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "tOGGLE cASE"
            Index           =   4
         End
      End
      Begin VB.Menu xxxmnuSeptF4 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuBullet 
         Caption         =   "Bullet Style"
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "Normal"
            Index           =   1
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "Numbered"
            Index           =   2
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "alpha lower case"
            Index           =   3
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "Alpha Upper Case"
            Index           =   4
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "roman lower case"
            Index           =   5
         End
         Begin VB.Menu mnuBulletStyle 
            Caption         =   "Roman Upper Case"
            Index           =   6
         End
      End
      Begin VB.Menu xxxmnuHighlight 
         Caption         =   "Highlight"
      End
      Begin VB.Menu xxxmnuF1 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuParagraph 
         Caption         =   "Paragraph..."
      End
      Begin VB.Menu xxxmnuTabs 
         Caption         =   "Tabs..."
         Shortcut        =   ^T
      End
      Begin VB.Menu xxxmnuClearFormat 
         Caption         =   "Clear Formatting"
      End
      Begin VB.Menu xxxmnuSepF8 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuWordwrap 
         Caption         =   "Wrap to Window"
         Index           =   0
      End
      Begin VB.Menu xxxmnuWordwrap 
         Caption         =   "Wrap to Ruler"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu xxxmnuWordwrap 
         Caption         =   "No Word Wrap"
         Index           =   2
      End
   End
   Begin VB.Menu xxxmnuHelpFiles 
      Caption         =   "Help"
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "iQ WordPad Manual"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "iQ Spell Check Help"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "About iQ WordPad"
         Index           =   3
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Begin VB.Menu mnuSpacing 
         Caption         =   "Normal Space   (Single Space)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSpacing 
         Caption         =   "One and a half (1.5 line spacing)"
         Index           =   1
      End
      Begin VB.Menu mnuSpacing 
         Caption         =   "Double Space   (2 line spacing)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMainText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'track mouse
'needed for xp theme
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Findwindow is used for email sending
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'For shortening path display in status bar
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As Long, ByVal lpszPath As String, ByVal dx As Long) As Long

'For ins, num and caps display in status bar
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

'Freeze rtb for updates when needed
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


'For ruler moving
Dim X1 As Single
Dim Y1 As Single
Dim Start As Boolean
Dim tCount As Integer

'for URL Linking
Private Const EM_CHARFROMPOS& = &HD7
Dim XX1 As Single
Dim YY1 As Single
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_USER = &H400
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
Private Const EM_AUTOURLDETECT = (WM_USER + 91)

'for bullet and number/special bullets
Dim BulletNormal As Boolean
Dim BulletNumber As Boolean

'Create instance of table
Dim RTFtable As clsRTFtable
Private Sub SaveAsWordDoc(mfile As String)
    Dim WordApp As Object
    Dim Document As Object
    On Error GoTo woops
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    Set Document = WordApp.Documents.Add
    Clipboard.Clear
    Clipboard.SetText Text1.TextRTF, vbCFRTF
    WordApp.ActiveDocument.Content.Paste
    Document.SaveAs mfile
    WordApp.Application.Quit
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    Exit Sub
woops:
    frmConvert.Hide
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Private Sub OpenWordDoc(mfile As String)
    Dim WordApp As Object
    On Error GoTo Errhandler
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open mfile
    WordApp.ActiveDocument.Content.Copy
    Text1.SelRTF = Clipboard.GetText(vbCFRTF)
    WordApp.Application.Quit
    Set WordApp = Nothing
    Screen.MousePointer = 0
    Exit Sub

Errhandler:
    frmConvert.Hide
    Set WordApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub


Public Sub EnableAutoURLDetection(RTB As RichTextBox)

    'enable auto URL detection
    
    SendMessage RTB.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    
'*************************************************************
' not using subclass method below because it interferes with other
' needed messages for richtext box.  Using RichWordOver function to get
' actual web address for shell execute.
'**************************************************************


    'subclass the parent of the RTB to receive EN_LINK notifications
    'Set FormSubClass = New clsSubClass
    'FormSubClass.Enable RTB.Parent.hwnd
    
    'set RTB to notify parent when user has clicked hyperlink
    'SendMessage RTB.hwnd, EM_SETEVENTMASK, 0&, ByVal ENM_LINK

End Sub
' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
Dim pt As POINTAPI
Dim pos As Integer
Dim start_pos As Integer
Dim end_pos As Integer
Dim ch As String
Dim txt As String
Dim txtlen As Integer

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
    
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "!" And ch <= "z") Or _
            ch = "_" _
        ) Then Exit For
        
    Next start_pos
    
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
    
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "!" And ch <= "z") Or _
            ch = "_" _
        ) Then Exit For
        
    Next end_pos
    
    end_pos = end_pos - 1

    If start_pos <= end_pos Then _
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
        
End Function
Private Sub CompactPath(FName As String)
   Dim lhDC As Long, lCtlWidth As Long
     
   Me.ScaleMode = vbPixels
   StatusBar1.ScaleMode = vbPixels
   lCtlWidth = lblIQFileName.Width - Me.DrawWidth
   lhDC = Me.hdc
    
   PathCompactPath lhDC, FName, lCtlWidth
    Me.ScaleMode = vbTwips
    StatusBar1.ScaleMode = vbTwips
    
   lblIQFileName.Caption = FName
   lblIQFileName.ToolTipText = IQFileName
   
End Sub


Private Sub InsertFile(TempFile As String)
    On Error Resume Next
    Dim FileNum As Integer
    Dim Temp As String
    
    FileNum = FreeFile

    Open TempFile For Binary As #FileNum
    
    Temp = String(LOF(FileNum), Chr$(0))
    
    Get #FileNum, , Temp
    
    Close #FileNum
    
    'check for Unicode text
    
    If Left(Temp, 2) = "ÿþ" Or Left(Temp, 2) = "þÿ" Then Temp = Replace(Right(Temp, Len(Temp) - 2), Chr(0), "")
    
    'now display
    txtInsert.TextRTF = Temp
    Temp = ""
End Sub

Private Sub RTFConvert()
On Error GoTo Errhandler

Text1.HideSelection = True
If Text1.SelLength = 0 Then
 Text1.SelStart = 0
 Text1.SelLength = Len(Text1.TextRTF)
End If
   ColorHighLite = QBColor(15)
   HighliteText (ColorHighLite)
txtInsert.Text = Text1.SelText
Text1.SelFontName = "Arial"
Text1.SelFontSize = 10
Text1.SelBold = False
Text1.SelText = txtInsert.Text
Text1.SelStart = 0
txtInsert.Text = ""
Text1.HideSelection = False
Text1.SetFocus
Exit Sub


Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 
End Sub
Private Sub CheckMSSpell()
Dim speller As Object
    On Error GoTo OpenError
    Set speller = CreateObject("Word.Application")
    
    Exit Sub
    
OpenError:
  xxxmnuSpellCheck.Enabled = False
  cmdToolButton(13).Enabled = False
End Sub


Private Sub CheckSpell()
'This is to check for Spell Check Anywhere utility
On Error GoTo Errhandler

There = Exist(App.Path & "\sa.exe")
If There Then
 'see if it's already running
 ret = IsAppRunning("sa.exe")
 If ret = False Then
  Shell App.Path & "\sa.exe", vbNormalNoFocus
 End If
Else
 xxxmnuSpellCheck.Enabled = False
 cmdToolButton(13).Enabled = False
End If

Exit Sub

Errhandler:
 xxxmnuSpellCheck.Enabled = False
 cmdToolButton(13).Enabled = False
 Exit Sub
End Sub


Private Sub ClearTabs()

If tCount = 1 Then
 Unload imgTab(1)
Else
 For i = 1 To tCount
  Unload imgTab(i)
 Next
End If

tCount = 0

End Sub


Private Sub FileOpen()
    On Error Resume Next
    Dim FileNum As Integer
    Dim Temp As String
    
    FileNum = FreeFile

    Open IQFileName For Binary As #FileNum
    
    Temp = String(LOF(FileNum), Chr$(0))
    
    Get #FileNum, , Temp
    
    Close #FileNum
    
    'check for Unicode text
    
    If Left(Temp, 2) = "ÿþ" Or Left(Temp, 2) = "þÿ" Then Temp = Replace(Right(Temp, Len(Temp) - 2), Chr(0), "")
    
    'now display
    Text1.TextRTF = Temp
    Temp = ""
    
End Sub
Private Sub EditMenuEnable()
    'Enable's/disables edit menu items and toolbar buttons
   On Error Resume Next
    Dim Enabled As Boolean
    Enabled = (Text1.SelLength > 0)
    xxxmnuCut.Enabled = Enabled
    xxxmnuCopy.Enabled = Enabled
    xxxmnuDelete.Enabled = Enabled
    xxxmnuSuperScript.Enabled = Enabled
    xxxmnuSubScript.Enabled = Enabled
    cmdToolButton(4).Enabled = Enabled
    cmdToolButton(5).Enabled = Enabled
    cmdToolButton(7).Enabled = Enabled
    
    xxxmnuUndo.Enabled = SendMessage(Text1.hwnd, EM_CANUNDO, 0&, 0&)
    cmdToolButton(8).Enabled = xxxmnuUndo.Enabled
 
    
    imgLeft.Move (Text1.SelIndent + 90) + Text1.SelHangingIndent
    imgHang.Move imgLeft.Left - Text1.SelHangingIndent
    imgRight.Move gMargin + 120 - (Text1.SelRightIndent)
    
    If Text1.SelTabCount > 0 Or tCount > 0 Then ShowTabs
    

End Sub


Private Sub PrintSelected()
 '---------------------------------------------------
 'routine to print highlighted, selected text only
 '---------------------------------------------------
 
 txtInsert.TextRTF = Text1.SelRTF
 PrintRTF txtInsert, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
 txtInsert.Text = ""
 
End Sub

Private Sub SentenceCase()
 Dim Temp As String
 Dim HoldText As String
 Dim SChar As String
 Dim charnum As Integer
 Dim flag As Boolean
 
 HoldText = Text1.SelText
 
 'this routine works only on lower case so be sure text is all lower
 HoldText = StrConv(HoldText, vbLowerCase)
  
 'Now parse the string to make sentence case
 For i = 1 To Len(HoldText)
  SChar = Mid$(HoldText, i, 1)
  charnum = Asc(SChar)
  
  'first letter always should be capital
  If i = 1 And charnum > 96 And charnum < 123 Then
   charnum = charnum - 32
   flag = True
  End If
 
 'Change to lower case if not first char in a sentence
 If flag = False And charnum > 64 And charnum < 91 Then
  charnum = charnum + 32
 End If
 
 'Change to uppercase if first char in a sentence
 If flag = True And charnum > 96 And charnum < 123 Then
  charnum = charnum - 32
  flag = False
 End If
 
 Temp = Temp & Chr$(charnum)
 
 'if character is period then next character will be capitalized
  If charnum = 46 Then flag = True
  
 'only for first character Flag should always be false
  If i = 1 Then flag = False
 
 
 Next i
 
 'Now assign to selected text
  Text1.SelText = Temp
 
End Sub

Private Sub SetFormatBar()
On Error Resume Next
 'Set Font and size
  cboFont.Text = Text1.SelFontName
  cboSize.Text = Int(Text1.SelFontSize + 0.5)
  
 'Set bold
  If Text1.SelBold = True Then
    butBold(0).Visible = False
    butBold(1).Visible = True
  Else
    butBold(0).Visible = True
    butBold(1).Visible = False
  End If
  
 'Set italic
  If Text1.SelItalic = True Then
    butItalic(0).Visible = False
    butItalic(1).Visible = True
  Else
    butItalic(0).Visible = True
    butItalic(1).Visible = False
  End If
  
 'set underline
  If Text1.SelUnderline = True Then
    butUnderline(0).Visible = False
    butUnderline(1).Visible = True
  Else
    butUnderline(0).Visible = True
    butUnderline(1).Visible = False
  End If
  
 'set alignment
  For i = 0 To 5 Step 2 'reset em all
   butAlign(i).Visible = True
   butAlign(i + 1).Visible = False
  Next
   'now set the active alignment
  If Text1.SelAlignment = rtfLeft Then
   butAlign(1).Visible = True
  ElseIf Text1.SelAlignment = rtfCenter Then
   butAlign(3).Visible = True
  ElseIf Text1.SelAlignment = rtfRight Then
   butAlign(5).Visible = True
  End If
  
  'Set bullet button
  If Text1.SelBullet = True Then
   BulletNormal = True
   butbullet(0).Visible = False
   butbullet(1).Visible = True
  Else
   BulletNormal = False
   butbullet(0).Visible = True
   butbullet(1).Visible = False
  End If
  
  'Set bullet number button
  If Text1.SelBullet = False And Text1.SelHangingIndent <> 0 Then
   BulletNumber = True
   butnumber(0).Visible = False
   butnumber(1).Visible = True
  Else
   BulletNumber = False
   butnumber(0).Visible = True
   butnumber(1).Visible = False
  End If
  
End Sub

Private Sub ShowTabs()
'clear previous

 On Error Resume Next
 ClearTabs
 
 For i = 0 To Text1.SelTabCount - 1
  tCount = tCount + 1
  Load imgTab(tCount)
  imgTab(tCount).Move Text1.SelTabs(i) + 150
  imgTab(tCount).Visible = True
 Next

End Sub

Private Sub ToggleCase()

 Dim Temp As String
 Dim SChar As String
 Dim charnum As Integer
 
 'Just reverse each character in selected text
 
 For i = 1 To Len(Text1.SelText)
  SChar = Mid$(Text1.SelText, i, 1)
  charnum = Asc(SChar)
   If charnum > 96 And charnum < 123 Then
    charnum = charnum - 32
   ElseIf charnum > 64 And charnum < 91 Then
    charnum = charnum + 32
   End If
  Temp = Temp & Chr$(charnum)
 Next
 Text1.SelText = Temp
End Sub

Private Sub UpdateLog()
 'this writes to any .LOG text files opened
 
 Text1.Text = Text1.Text & Format$(Now, "h:mm AMPM m/dd/yyyy") & vbCrLf
 'Text1.SelStart = Len(Text1.Text) + 1 '-uncomment this line to scroll to end of doc
 
End Sub


Private Sub butAlign_Click(Index As Integer)
'0=Alignleft off 1=Alignleft on
'2=Aligncntr off 3=aligncntr on
'4=alignright off 5=alignright on
For i = 0 To 5 Step 2
 butAlign(i).Visible = True
 butAlign(i + 1).Visible = False
Next

 Select Case Index
 
  Case 0
   butAlign(1).Visible = True
   Text1.SelAlignment = rtfLeft
  
  Case 2
   butAlign(3).Visible = True
   Text1.SelAlignment = rtfCenter
   
  Case 4
   butAlign(5).Visible = True
   Text1.SelAlignment = rtfRight
   
  Case Else
   butAlign(Index).Visible = True
   
 End Select
   
 Text1.SetFocus
End Sub

Private Sub butBold_Click(Index As Integer)

Select Case Index
 Case 0
  butBold(0).Visible = False
  butBold(1).Visible = True
  
 Case 1
  butBold(0).Visible = True
  butBold(1).Visible = False
 End Select
 
 Text1.SelBold = Not Text1.SelBold
 Text1.SetFocus
 
End Sub

Private Sub butbullet_Click(Index As Integer)
 
 Select Case Index
  
  Case 0
   mnuBulletStyle_Click (1)
  
  Case 1
   mnuBulletStyle_Click (0)
 
 End Select

End Sub

Private Sub butCaseChange_Click()
 PopupMenu mnuChangeCase
 picFormatbar.Refresh
End Sub

Private Sub butColor_Click()
 On Error GoTo Errhandler
 HighColor = False
 frmColorPick.Show 1
 Exit Sub
 
Errhandler:
 
 picFormatbar.Refresh
 RulerBar.Refresh
  picRuler.Refresh
 picFormatbar.Refresh
 Text1.SetFocus
 
 Exit Sub

End Sub

Private Sub butHighlight_Click()
 On Error GoTo Errhandler
 
 HighColor = True
 frmColorPick.Show 1
 Exit Sub
 
Errhandler:
 
 picFormatbar.Refresh
 Text1.SetFocus
 
 Exit Sub
 
 
End Sub

Private Sub butItalic_Click(Index As Integer)

Select Case Index
 Case 0
  butItalic(0).Visible = False
  butItalic(1).Visible = True
  
 Case 1
  butItalic(0).Visible = True
  butItalic(1).Visible = False
 End Select
 
 Text1.SelItalic = Not Text1.SelItalic
 Text1.SetFocus
End Sub

Private Sub butnumber_Click(Index As Integer)
 
 Select Case Index
 
  Case 0
   mnuBulletStyle_Click (2)
  
  Case 1
   mnuBulletStyle_Click (0)
   
  Case 2
   butnumber(3).Visible = True
   PopupMenu xxxmnuBullet
   picFormatbar.Refresh
   butnumber(3).Visible = False
 End Select

 
End Sub

Private Sub butRecentFonts_Click()
 PopupMenu xxxmnuRecentFonts
 picToolbar.Refresh
 Text1.SetFocus
End Sub

Private Sub butReferenes_Click()
 PopupMenu mnuRefer
 picToolbar.Refresh
 
End Sub

Private Sub butSpacing_Click()
 For i = 0 To 2
  mnuSpacing(i).Checked = False
 Next
 
 mnuSpacing(LineSpacing).Checked = True
 
 PopupMenu mnuHidden
 picFormatbar.Refresh
 
 Text1.SetFocus
End Sub

Private Sub butUnderline_Click(Index As Integer)
Select Case Index
 Case 0
  butUnderline(0).Visible = False
  butUnderline(1).Visible = True
  
 Case 1
  butUnderline(0).Visible = True
  butUnderline(1).Visible = False
 End Select
 
 Text1.SelUnderline = Not Text1.SelUnderline
 Text1.SetFocus
End Sub

Private Sub cboFont_Click()
   Text1.SelFontName = cboFont.Text
   Text1.SetFocus
   UpDateFontMenu Text1.SelFontName
End Sub


Private Sub cboSize_Click()
    Text1.SelFontSize = Val(cboSize.Text)
    Text1.SetFocus
End Sub


Private Sub cmdToolButton_Click(Index As Integer)

 Select Case Index
 
  Case 0 'New
   xxxmnuNew_Click
  
  Case 1 'Open
   xxxmnuOpen_Click
  
  Case 2 'Save
   xxxmnuSave_Click
  
  Case 3 'Print
   xxxmnuPrint_Click
  
  Case 4 'Cut
   SendMessage Text1.hwnd, WM_CUT, 0&, 0&
  
  Case 5 'Copy
   SendMessage Text1.hwnd, WM_COPY, 0&, 0&
  
  Case 6 'Paste
   xxxmnuPaste_Click
   
  Case 7 'Delete
   SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
    
  Case 8 'Undo
   SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
   EditMenuEnable
   Text1.SetFocus
  
  Case 9 'Find/replace
   xxxmnuReplace_Click

  Case 10 'insert picture
   xxxmnuInsertPic_Click
   
  Case 11 'insert date and time
  
   frmDateTime.Show 1
   
  Case 12 ' Print preview
   xxxmnuPrintPreview_Click
     
  Case 13 ' Spellcheck
   xxxmnuSpellCheck_Click
   
  Case 14 ' Symbol Insert
   xxxmnuCharMap_Click
   
End Select



End Sub

Private Sub Form_Activate()
 On Error Resume Next
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub Form_Initialize()
'  InitCommonControls
InitCommonControlsXP
End Sub

Private Sub Form_Load()
 On Error Resume Next
 Me.Top = (Screen.Height - Me.Height) / 2.4
 Me.Left = (Screen.Width - Me.Width) / 2
 Me.BackColor = Text1.BackColor
 mnuHidden.Visible = False
 
 'Create table from class
 Set RTFtable = New clsRTFtable
 
 'Enable RTB to detect Web URL's
 EnableAutoURLDetection Text1
 
 'Load fonts and sizes for Format Bar
 
    For i = 1 To Screen.FontCount - 1
        If Screen.Fonts(i) <> "" Then cboFont.AddItem Screen.Fonts(i)
    Next i
    
    For i = 8 To 12 Step 1
        cboSize.AddItem i
    Next
    
    For i = 14 To 28 Step 2
        cboSize.AddItem i
    Next
    
    cboSize.AddItem 36
    cboSize.AddItem 48
    cboSize.AddItem 72
    cboFont.Text = "Arial"
    cboSize.Text = "10"
    NameFont = "Arial"
    SizeFont = 10
    BoldFont = True
    ItalicFont = False
    ColorFont = 0
    
    Text1.Font.Name = cboFont.Text
    Text1.Font.Size = SizeFont
    Text1.SelFontName = cboFont.Text
    Text1.SelFontSize = SizeFont
    Text1.SelBold = BoldFont
    
 'This will extend width and height of font combo dropdown
 SubClass cboFont.hwnd, True, 200, 300, 0
 
'Set Formatbar buttons
 butBold(1).Visible = True
 butItalic(1).Visible = False
 butUnderline(1).Visible = False
 
 For i = 1 To 5 Step 2
  butAlign(i).Visible = False
 Next
  butAlign(0).Visible = False
  butAlign(1).Visible = True
  butbullet(1).Visible = False
  butnumber(1).Visible = False
  
 'Set gripper location
 imgGripper.Top = 150
 imgGripper.Left = 4220

'initialise startup
ColorHighLite = 65535 'yellow
ToolbarOn = True
FormatbarOn = True
RulerOn = True
iQMeasurement = 0
LineSpacing = 0
StatusbarOn = True
WrapOn = 1
IQFileName = "Untitled"
frmMainText.Caption = " " & IQFileName & " - " & "iQ WordPad"
cmDlg.Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Other Text (*.bas;*.bat;*.csv;*.dat;*.ini;*.lst)|*.bas;*.bat;*.csv;*.dat;*.ini;*.lst|All Files (*.*)|*.*"

GetRecentFiles 'from ini file
GetOptions 'from ini file
GetRecentFonts 'from ini file

If RecentFonts(1) = "" Then
 UpDateFontMenu Text1.Font.Name
End If

'printer defaults
gLeft = 1440
gRight = 1440
gTop = 1440
gBottom = 1440
gPaperSize = 1 'letter
gPaperWidth = 8.5
gPaperHeight = 11
gOrientation = 1 'portrait
gHeader = ""
gFooter = ""

'set ruler size
If iQMeasurement = 1 Then
 xxxmnuRulerOption(0).Checked = False
 xxxmnuRulerOption(1).Checked = True
 picRuler.Picture = imgRuler(1).Picture
End If

GetPrintOptions 'from ini file

If gOrientation = 1 Then
 gMargin = (gPaperWidth * 1440) - (gLeft + gRight)
Else
 gMargin = (gPaperHeight * 1440) - (gLeft + gRight)
End If

imgMargin.Move gMargin + 150
imgRight.Move gMargin + 90
Text1.RightMargin = gMargin + 90
EditMenuEnable
 
'***check for auto load from command line********************************
If Len(Command$) Then
 IQFileName = Command$
 If Left$(IQFileName, 1) = Chr$(34) Then 'some programs put quotes on command line.
  IQFileName = Mid$(IQFileName, 2, Len(IQFileName) - 2)
 End If
 
 There = FileExists(IQFileName)
  If There Then
   If LCase(Right$(IQFileName, 3)) = "rtf" Then
    Text1.LoadFile IQFileName, rtfRTF
   ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    frmConvert.Show
    DoEvents
    OpenWordDoc (IQFileName)
    frmConvert.Hide
   Else
    FileOpen
   End If
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
  End If
End If

'***************************************************
'check for spell check anywhere
 'CheckSpell
'check for MS Word Spell Checker
 'CheckMSSpell
'*******************************************************

CompactPath (LCase(IQFileName))
xxxmnuWordwrap_Click (WrapOn)
 
 'XPNetMenu1.MainBarGradientDire = usVerticality
' XPNetMenu1.CheckBack = 1743087
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Text1.DataChanged = True And Len(Text1.Text) Then
 
  ret = MsgBox("File has changed. Do you wish to save " & IQFileName & "?", vbQuestion + vbYesNoCancel, "iQ WordPad - Save?")
  
  If ret = vbYes Then
     If IQFileName = "Untitled" Then
      xxxmnuSaveAs_Click
     Else
      xxxmnuSave_Click
     End If
   ElseIf ret = vbCancel Then
    Cancel = True
    Exit Sub
  End If
 End If
 
 Text1.Text = ""
  
 WriteOptions

For i = 1 To tCount
 Set imgTab(i) = Nothing
 Unload imgTab(i)
Next

For i = Forms.Count - 1 To 0 Step -1
 Unload Forms(i)
Next

''see if spell check is running and if so close it
 'ret = IsAppRunning("sa.exe")
 'If ret = True Then
    'Call EnumWindows(AddressOf EnumWindowProc, &H0)
    'ret = SendMessage(ClosehWnd, WM_CLOSE, 0, 0)
 'End If

End Sub

Private Sub Form_Resize()
 Dim AlignHeight As Integer
 If Me.WindowState = 1 Then Exit Sub
 If Me.Height < 3000 Then
  Me.Height = 3000
  Me.Enabled = False
  Me.Enabled = True
 End If
 
 If Me.Width < 3000 Then
  Me.Width = 3000
  Me.Enabled = False
  Me.Enabled = True
 End If
 

 Text1.Width = Me.Width - 450
 
'------------------
'lets investigate
'------------------
Dim theight As Integer

'check for toolbars

  If ToolbarOn Then
   theight = theight + picToolbar.Height
   AlignHeight = picToolbar.Height
  End If

  If FormatbarOn Then
   theight = theight + picFormatbar.Height
   picFormatbar.Top = AlignHeight - 10
   AlignHeight = AlignHeight + picFormatbar.Height
  End If
  
  If RulerOn Then
   theight = theight + RulerBar.Height
   RulerBar.Top = AlignHeight - 10
  End If
    
  Text1.Top = theight

 
'Adjust for Text Height status
 If StatusbarOn = True Then
  theight = StatusBar1.Height + 820
 Else
  theight = 820
 End If
 
  If ToolbarOn Then theight = theight + picToolbar.Height
  If RulerOn Then theight = theight + RulerBar.Height
  If FormatbarOn Then theight = theight + picFormatbar.Height
  Text1.Height = Me.Height - theight
  picLnCol.Left = Me.Width - 4585


End Sub









Private Sub imgGripper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Negate VB's call to SetCapture, and tell Windows
   ' that the user is trying to resize the form.
   ReleaseCapture
   SendMessage hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub


Private Sub imgHang_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 xxxmnuParagraph_Click
 Exit Sub
End Sub


Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 X1 = X      'This holds the previouse X value for the mouse..only for moving
 Y1 = Y      'This holds the previous y value for the mouse..only for moving
 Start = True
End Sub


Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo Hell:
X2 = X
Y2 = Y

If (Start) And Button = 1 Then

 With imgLeft
  If .Left - X1 + X2 < 90 Then
   .Left = 90
  Else
   .Move imgLeft.Left - X1 + X2
  End If
 End With
 
 imgHang.Left = imgLeft.Left
 
End If


Hell:
 Exit Sub
End Sub


Private Sub imgLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Start = False
 Text1.SelIndent = imgLeft.Left - 90
End Sub


Private Sub imgMargin_Click()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 X1 = X      'This holds the previouse X value for the mouse..only for moving
 Y1 = Y      'This holds the previous y value for the mouse..only for moving
 Start = True
End Sub


Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo Hell:
X2 = X
Y2 = Y

If (Start) And Button = 1 Then

 With imgRight
  If .Left - X1 + X2 > imgMargin.Left Then
    imgRight.Left = imgMargin.Left
  Else
   .Move imgRight.Left - X1 + X2
  End If
 End With

End If


Hell:
Exit Sub
End Sub


Private Sub imgRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Start = False
 Text1.SelRightIndent = gMargin + 190 - (imgRight.Left + 90)
End Sub


Private Sub imgTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  frmTabs.Show 1
  ShowTabs
  Exit Sub
 End If
 
 X1 = X      'This holds the previouse X value for the mouse..only for moving
 Y1 = Y      'This holds the previous y value for the mouse..only for moving
 Start = True
End Sub


Private Sub imgTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo Hell:
X2 = X
Y2 = Y

If (Start) And Button = 1 Then

 With imgTab(Index)
  If .Left - X1 + X2 < 150 Then
   .Left = 150
  Else
   .Move .Left - X1 + X2
  End If
 End With
 
End If

Hell:
 Exit Sub
End Sub


Private Sub imgTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Start = False
 Text1.SelTabs(Index - 1) = imgTab(Index).Left - 150
End Sub

Private Sub mnuAttributes_Click(Index As Integer)

 Select Case Index
 
  Case 0 'text foreground color
   butColor_Click
   Exit Sub

  Case 1 'highlight color
   HighColor = True
   frmColorPick.Show 1
   Exit Sub
   
   
  Case 2 'Bold
   Text1.SelBold = Not Text1.SelBold
  
  Case 3 'Italic
   Text1.SelItalic = Not Text1.SelItalic
   
  Case 4 'Underline
   Text1.SelUnderline = Not Text1.SelUnderline
   
  Case 5 'Strikethrough
   Text1.SelStrikeThru = Not Text1.SelStrikeThru
   
 End Select
 
 
Errhandler:
 If FormatbarOn = True Then SetFormatBar
 Exit Sub
End Sub

Private Sub mnuBulletStyle_Click(Index As Integer)
'reset all bullets
 Text1.SelBullet = False
 BulletNormal = False
 BulletNumber = False
 Text1.SetFocus
 
'now set bullet choice
 Select Case Index
 
  Case 0
   Text1.SelBullet = False
   BulletNormal = False
   DoEvents
   
  Case 1
   BulletNormal = True
   SendKeys "^+L", True
   DoEvents
   
  Case 2
   Text1.SetFocus
   BulletNumber = True
   SendKeys "^+L", True
    DoEvents
   SendKeys "^+L", True
    DoEvents
  
  Case Else
    
    For i = 1 To Index
     SendKeys "^+L", True
     DoEvents
    Next

  
 End Select

SetFormatBar
EditMenuEnable
butnumber(3).Visible = False

End Sub

Private Sub mnuCalc_Click()
 On Error Resume Next
 
 Shell "calc.exe", vbNormalFocus
 
End Sub



Private Sub mnuCaseChange_Click(Index As Integer)
    
 Select Case Index
 
  Case 0 ' Sentence case
   SentenceCase
 
  Case 1 ' lower case
   Text1.SelText = StrConv(Text1.SelText, vbLowerCase)
   
  Case 2 ' upper case
   Text1.SelText = StrConv(Text1.SelText, vbUpperCase)

  Case 3 ' proper case
   Text1.SelText = StrConv(Text1.SelText, vbProperCase)
   
  Case 4 ' tOGGLE cASE
   ToggleCase

 End Select
   
     Text1.SetFocus
     
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
 
 If RecentDocs(Index) = IQFileName Then Exit Sub 'trying to open current file
 
  If Text1.DataChanged And Len(Text1.Text) Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "iQ WordPad - Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 On Error Resume Next
 Screen.MousePointer = 11
 IQFileName = RecentDocs(Index)
 
 There = FileExists(IQFileName)
 If There Then
  Text1.Text = ""
   If LCase(Right$(IQFileName, 3)) = "rtf" Then
    Text1.LoadFile IQFileName, rtfRTF
   ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    picToolbar.Refresh
    frmConvert.Show
    DoEvents
    OpenWordDoc (IQFileName)
    frmConvert.Hide
   Else
    FileOpen
   End If
  ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
  LineSpacing = 0
  frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
  CompactPath (LCase(IQFileName))
  Text1.DataChanged = False
  UpDateFileMenu IQFileName
  
  'is this a log file
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   
    Text1.SetFocus
    picToolbar.Refresh
    RulerBar.Refresh
    picRuler.Refresh
    picFormatbar.Refresh
    If FormatbarOn = True Then SetFormatBar
    ShowTabs
    Screen.MousePointer = 0
    Exit Sub
 
  Else
 
    MsgBox IQFileName & " no longer available!", vbOKOnly, "File Not Found"
    If FormatbarOn = True Then SetFormatBar
    Screen.MousePointer = 0
    Exit Sub
    
 End If
 
End Sub

Private Sub mnuRecentFonts_Click(Index As Integer)
 
 If Index > 0 Then
  Text1.SelFontName = mnuRecentFonts(Index).Caption
  cboFont.Text = Text1.SelFontName
 End If
  
End Sub

Private Sub mnuReferences_Click(Index As Integer)
 Dim Ref As String
 
 On Error Resume Next
 
 Select Case Index
 
  Case 0 'Dictionary
   Ref = "http://www.m-w.com/dictionary"
   If Text1.SelLength > 0 Then Ref = Ref & "/" & Text1.SelText
   
    
  Case 1 'Thesaurus
   Ref = "http://thesaurus.reference.com"
   If Text1.SelLength > 0 Then Ref = Ref & "/browse/" & Text1.SelText
   
  Case 2 'Wikipedia
   Ref = "http://en.wikipedia.org"
   If Text1.SelLength > 0 Then Ref = Ref & "/wiki/" & Text1.SelText
   
  Case 3 'Google
  Ref = "http://www.google.com/"
   If Text1.SelLength > 0 Then Ref = Ref & "search?hl=en&q=" & Text1.SelText
   
 End Select
 
 ret = ShellExecute(0&, vbNullString, Ref, vbNullString, vbNullString, vbNormalFocus)
 
End Sub

Private Sub mnuSpacing_Click(Index As Integer)
 LineSpacing = Index
 SelLineSpacing Text1, LineSpacing
End Sub

Private Sub picFormatbar_Click()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub picRuler_Click()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub picRuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Button = 1 Then
  If IsNull(Text1.SelTabCount) Then
   ClearTabs
   Text1.SelTabCount = 0
  End If
  
  tCount = tCount + 1
  Text1.SelTabCount = Text1.SelTabCount + 1
  Text1.SelTabs(Text1.SelTabCount - 1) = X - 150
  Load imgTab(tCount)
  imgTab(tCount).Move Text1.SelTabs(Text1.SelTabCount - 1) + 150
  imgTab(tCount).Visible = True
 End If

End Sub


Private Sub picToolbar_Click()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub RulerBar_Click()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub Text1_Change()
 
 Text1.DataChanged = True
 

'only accessed if user pressed Ctrl-V paste key
If PasteKey = True And Clipboard.GetFormat(vbCFRTF) Then
  SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
  SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
  Text1.SelRTF = Clipboard.GetText(vbCFRTF)
  DoEvents
  PasteKey = False
End If

 
End Sub

Private Sub Text1_GotFocus()
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

'needed for send mail
 If KeyCode = vbKeyI And Shift = 4 Then 'alt i
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

If KeyCode = vbKeyV And Shift = 2 Then
 PasteKey = True
End If


 If KeyCode = vbKeyH And Shift = 2 Then
  xxxmnuReplace_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

 If KeyCode = vbKeyF And Shift = 2 Then
  xxxmnuFind_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

 If KeyCode = vbKeyG And Shift = 2 Then
  xxxmnuGoto_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

If KeyCode = vbKeyF5 Then
 xxxmnuTimeDate_Click
 KeyCode = 0
 Exit Sub
End If
 
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
 If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 XX1 = X
 YY1 = Y
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    
  If Button = vbRightButton Then
     PopupMenu xxxmnuEdit
     Exit Sub
  End If

'This is for Internet URL linking
  If Shift = 2 And Button = vbLeftButton Then 'ctrl key and left click
  
   On Error Resume Next
   Dim txt As String

   txt = RichWordOver(Text1, XX1, YY1)
   
   txt = LCase(txt)
    
   'sometimes rtb putting <> characters at front and back. If there parse out
   If Left$(txt, 5) = "<www." Or Left$(txt, 8) = "<http://" Then
    txt = Mid$(txt, 2, Len(txt) - 2)
   End If
   
   'if this is a URL then start associated browser and goto web site
   If Left$(txt, 4) = "www." Or Left$(txt, 7) = "http://" Then
    Screen.MousePointer = 11
    ret = ShellExecute(0&, vbNullString, txt, vbNullString, vbNullString, vbNormalFocus)
    Screen.MousePointer = 0
    Exit Sub
   End If
  
   'is it email
   If InStr(1, txt, "@") Then
    If LCase(Right(txt, 4)) = ".com" Or LCase(Right(txt, 4)) = ".net" Then
     ' Start email
     txt = "mailto:" & txt
     ShellExecute Me.hwnd, vbNullString, txt, vbNullString, "C:\", vbNormalFocus
    End If
   End If
   
 End If
'----------end of URL check

End Sub


Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'for unexpected errors
On Error GoTo Errhandler


If Data.GetFormat(vbCFFiles) Then 'legit file to drop?
 'do we need to save current doc?
 If Text1.DataChanged Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "iQ WordPad - Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 Dim OLEFileName As String
 OLEFileName = Data.Files.Item(1)
 
 There = FileExists(OLEFileName) 'just be sure it's there
  If There Then
   IQFileName = OLEFileName
   If LCase(Right$(IQFileName, 3)) = "rtf" Then
    Text1.LoadFile IQFileName, rtfRTF
   ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    picToolbar.Refresh
    frmConvert.Show
    DoEvents
    OpenWordDoc (IQFileName)
    frmConvert.Hide
   Else
    FileOpen
   End If
   ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   LineSpacing = 0
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
  Else
   MsgBox OLEFileName & " missing or invalid!", vbOKOnly + vbCritical, "iQ WordPad Drag/Drop Error"
   Exit Sub
  End If

End If

CompactPath (LCase(IQFileName))
Text1.SetFocus
Exit Sub

Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"

End Sub

Private Sub Text1_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub


Private Sub Text1_SelChange()
 
 EditMenuEnable
 
 If FormatbarOn = True Then SetFormatBar
 
 
End Sub

Private Sub tmrClipboard_Timer()
 With Clipboard
  If .GetFormat(1) = True Or .GetFormat(2) = True Or .GetFormat(3) = True Or .GetFormat(8) = True Then
    xxxmnuPaste.Enabled = True
  Else
    xxxmnuPaste.Enabled = False
  End If
 End With
 
    cmdToolButton(6).Enabled = xxxmnuPaste.Enabled
    
  If StatusbarOn = True Then
  
    Dim b(0 To 254) As Byte

    GetKeyboardState b(0)
    If b(vbKeyNumlock) Then
       lblInsNumCap(2).Visible = True
    Else
       lblInsNumCap(2).Visible = False
    End If
    
    If b(vbKeyCapital) Then
      lblInsNumCap(1).Visible = True
    Else
       lblInsNumCap(1).Visible = False
    End If
    
    If b(vbKeyInsert) Then
      lblInsNumCap(0).Visible = False
      Else
      lblInsNumCap(0).Visible = True
    End If
  End If

End Sub

Private Sub XPNetMenu1_CustomDrawItemFont(Font As stdole.Font, Caption As String, ForeColour As stdole.OLE_COLOR)
  On Error GoTo Errhandler
 
 If Caption = "    ---  Recent Fonts List  ---           " Then
  Font.Bold = True
  'Font.Size = Font.Size + 1
  Font.Underline = True
  ForeColour = QBColor(12)
 End If
 
 For i = 1 To 8
  If Caption = mnuRecentFonts(i).Caption Then
   Font.Name = Caption
   Font.Size = Font.Size + 1
  End If
 Next

  
Errhandler:
 Exit Sub
End Sub

Private Sub xxxmnuBullet_Click()
'    With Text1
'        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
'            .SelBullet = True
'            .SelIndent = 0.5
'            .SelHangingIndent = -1.5
'        ElseIf .SelBullet = True Then
'            .SelBullet = False
'            .SelHangingIndent = False
'        End If
'     SetFormatBar
'    End With
End Sub

Private Sub xxxmnuCharMap_Click()
On Error GoTo Errhandler

 'display iQ charmap window
 Screen.MousePointer = 11
 frmSymbols.Show 1

 Exit Sub
 

Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 'MsgBox "Could not find charmap.exe, character map program.", vbOKOnly, "iq Error Information"
 
 Exit Sub
 
End Sub

Private Sub xxxmnuClearFormat_Click()
  
  ret = MsgBox("This will remove all formatting includng images. (This may take a few seconds)  Do you wish to continue?", vbYesNoCancel + vbQuestion, "iQ Notepad Query")
  If ret = vbCancel Or ret = vbNo Then Exit Sub
  If ret = vbYes Then
   Screen.MousePointer = 11
   Call RTFConvert
   Screen.MousePointer = 0
   Exit Sub
  End If

End Sub

Private Sub xxxmnuCopy_Click()
 
 SendMessage Text1.hwnd, WM_COPY, 0&, 0& 'Copy

End Sub

Private Sub xxxmnuCut_Click()

 SendMessage Text1.hwnd, WM_CUT, 0&, 0&
 
End Sub



Private Sub xxxmnuDelete_Click()

 SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
 
End Sub

Private Sub xxxmnuExit_Click()

 Unload Me
 
End Sub

Private Sub xxxmnuFind_Click()

 'Show Find API dialog
  Dim s As String
  If Text1.SelLength > 0 Then s = Text1.SelText Else s = ""
  ShowFind Me, Text1, FR_DOWN, s
  
End Sub

Private Sub xxxmnuFont_Click()
On Error GoTo Errhandler

 cmDlg.flags = cdlCFEffects Or cdlCFForceFontExist Or cdlCFScreenFonts
 
 cmDlg.FontName = Text1.SelFontName
 cmDlg.FontSize = Text1.SelFontSize
 cmDlg.Color = Text1.SelColor
 cmDlg.FontBold = Text1.SelBold
 cmDlg.FontItalic = Text1.SelItalic
 cmDlg.FontUnderline = Text1.SelUnderline
 cmDlg.FontStrikethru = Text1.SelStrikeThru
 
GetFont:
 
 cmDlg.ShowFont
 
      With cmDlg
        Text1.SelFontName = .FontName
        Text1.SelFontSize = .FontSize
        Text1.SelBold = .FontBold
        Text1.SelItalic = .FontItalic
        Text1.SelColor = .Color
        Text1.SelStrikeThru = .FontStrikethru
        Text1.SelUnderline = .FontUnderline
        cboFont.Text = .FontName
        UpDateFontMenu .FontName
      End With
      

  Exit Sub
      
Errhandler:
 If Err = 94 Then Resume GetFont 'Null error
 Exit Sub
 
End Sub

Private Sub xxxmnuFormatBar_Click()
 If FormatbarOn = False Then
  FormatbarOn = True
  xxxmnuFormatBar.Checked = True
  picFormatbar.Visible = True
  Form_Resize
  Text1.SetFocus
  SetFormatBar
 Else
  FormatbarOn = False
  xxxmnuFormatBar.Checked = False
  picFormatbar.Visible = False
  Form_Resize
 End If
 
End Sub

Private Sub xxxmnuGoto_Click()

 frmGoto.Show 1
 
End Sub

Private Sub xxxmnuHelp_Click(Index As Integer)

 On Error Resume Next
 Dim ftemp As String
 

 Select Case Index
 
  Case 0 'iQ WordPad PDF Manual if avail.
   ftemp = App.Path & "\iqwordpad.pdf"
   There = FileExists(ftemp)
   If Not There Then
    MsgBox "iQ WordPad PDF help manual not found.  Contact iqProPlus", vbOKOnly + vbInformation, "iQ WordPad Help"
    Exit Sub
   End If
   
   ret = ShellExecute(0&, vbNullString, ftemp, vbNullString, vbNullString, vbNormalFocus)
   Exit Sub
   
  Case 1 'Spell Check Help
  
   ftemp = App.Path & "\wspelldlg.chm"
   There = FileExists(ftemp)
   If Not There Then
    MsgBox "iQ WordPad Spell Check Help not found.  Contact iqProPlus", vbOKOnly + vbInformation, "iQ WordPad Help"
    Exit Sub
   End If
   
   ret = ShellExecute(0&, vbNullString, ftemp, vbNullString, vbNullString, vbNormalFocus)
   Exit Sub
   
  Case 3
  
   frmAbout.Show 1
   
 End Select
  
End Sub

Private Sub xxxmnuHighlight_Click()
 
 On Error GoTo Errhandler
 
 HighColor = True
 frmColorPick.Show 1
 Exit Sub
 
Errhandler:
 
 picFormatbar.Refresh
 Text1.SetFocus
 
 Exit Sub
    
End Sub

Private Sub xxxmnuInsDocument_Click()
 On Error GoTo Errhandler
 Dim TempFile As String
 
 cmDlg.DialogTitle = "Insert/Merge Document"
 cmDlg.FileName = ""
 cmDlg.Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Other Text (*.bas;*.bat;*.csv;*.dat;*.ini;*.lst)|*.bas;*.bat;*.csv;*.dat;*.ini;*.lst|All Files (*.*)|*.*"
 cmDlg.FilterIndex = 1
 cmDlg.ShowOpen
 
 Screen.MousePointer = 11
 
 TempFile = cmDlg.FileName
 
 
 There = FileExists(TempFile)
 If There Then
  txtInsert.Text = ""
   If LCase(Right$(TempFile, 3)) = "rtf" Then
    txtInsert.LoadFile TempFile, rtfRTF
   Else
    InsertFile (TempFile)
   End If
   
   txtInsert.SelStart = 0
   txtInsert.SelLength = Len(txtInsert.SelRTF)
   Text1.SelRTF = txtInsert.TextRTF
   DoEvents
   Text1.SetFocus
 Else
   MsgBox TempFile & " Not Found!", vbOKOnly, "File Not Found"
 End If

picToolbar.Refresh
RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
DoEvents
If FormatbarOn = True Then SetFormatBar
Screen.MousePointer = 0
Exit Sub
 
Errhandler:
 If FormatbarOn = True Then SetFormatBar
 RulerBar.Refresh
 Screen.MousePointer = 0
 Exit Sub
End Sub

Private Sub xxxmnuInsertPic_Click()
On Error GoTo Errhandler

cmDlg.DialogTitle = "Insert Picture"
cmDlg.Filter = "All Supported Pictures (*.bmp;*.dib;*.gif;*.jpg;*.ico;*.wmf)|*.bmp;*.dib;*.gif;*.jpg;*.ico;*.wmf|Bitmap (*.bmp;*.dib)|*.bmp;*.dib|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|Icon (*.ico)|*.ico|MetaFile (*.wmf)|*.wmf"
cmDlg.FilterIndex = 1
cmDlg.ShowOpen
    
    Screen.MousePointer = 11
    
    If cmDlg.FileName <> "" Then
        picInsert.Picture = LoadPicture(cmDlg.FileName)
        picInsert.Picture = picInsert.Image 'this makes icon stdpicture object
        DoEvents
        Clipboard.Clear
        Clipboard.SetData picInsert.Picture
        SendMessage Text1.hwnd, WM_PASTE, 0, 0
    End If
    
    Screen.MousePointer = 0
  
 Exit Sub
 
Errhandler:
     Screen.MousePointer = 0
     If Err <> 32755 Then '32755 is cancel error
     
       MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Picture Error!"
       
     End If

End Sub

Private Sub xxxmnuNew_Click()
 If Text1.DataChanged And Len(Text1.Text) Then
 
  ret = MsgBox("Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "iQ WordPad - Save?")
  
  If ret = vbCancel Then Exit Sub
  
  If ret = vbNo Then
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
   Text1.DataChanged = False
   Text1.SetFocus
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    CompactPath (LCase(IQFileName))
   End If
   Text1.SelFontName = NameFont
   Text1.SelFontSize = SizeFont
   Text1.SelBold = BoldFont
   Text1.SelItalic = ItalicFont
   Text1.SelColor = ColorFont
   LineSpacing = 0
   If FormatbarOn = True Then SetFormatBar
   ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   picToolbar.Refresh
   RulerBar.Refresh
   picRuler.Refresh
   picFormatbar.Refresh
   If FormatbarOn = True Then SetFormatBar
   Exit Sub
  End If
  
  If IQFileName = "Untitled" Then
   xxxmnuSaveAs_Click
   If fCancel = True Then Exit Sub
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & IQFileName & " - " & "iQ WordPad"
   Text1.DataChanged = False
   Text1.SetFocus
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    CompactPath (LCase(IQFileName))
   End If
   Text1.SelFontName = NameFont
   Text1.SelFontSize = SizeFont
   Text1.SelBold = BoldFont
   Text1.SelItalic = ItalicFont
   Text1.SelColor = ColorFont
   If FormatbarOn = True Then SetFormatBar
   ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   picToolbar.Refresh
   RulerBar.Refresh
   picRuler.Refresh
   picFormatbar.Refresh
   If FormatbarOn = True Then SetFormatBar
   Exit Sub
  Else
   xxxmnuSave_Click
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & IQFileName & " - " & "iQ WordPad"
   Text1.DataChanged = False
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    CompactPath (LCase(IQFileName))
   End If
   Text1.SetFocus
   Text1.SelFontName = NameFont
   Text1.SelFontSize = SizeFont
   Text1.SelBold = BoldFont
   Text1.SelItalic = ItalicFont
   Text1.SelColor = ColorFont
   If FormatbarOn = True Then SetFormatBar
   ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   picToolbar.Refresh
   RulerBar.Refresh
   picRuler.Refresh
   picFormatbar.Refresh
   If FormatbarOn = True Then SetFormatBar
   Exit Sub
  End If
 
 Else 'it's not dirty so clear out text and reset FileName
  
  Text1.Text = ""
  Text1.DataChanged = False
  IQFileName = "Untitled"
  frmMainText.Caption = " " & IQFileName & " - " & "iQ WordPad"
  Text1.SetFocus
  If StatusbarOn = True Then
   lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
   CompactPath (LCase(IQFileName))
  End If
   Text1.SelFontName = NameFont
   Text1.SelFontSize = SizeFont
   Text1.SelBold = BoldFont
   Text1.SelItalic = ItalicFont
   Text1.SelColor = ColorFont
 End If
   ClearTabs
   Text1.SetFocus
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   picToolbar.Refresh
   RulerBar.Refresh
   picRuler.Refresh
   picFormatbar.Refresh
   If FormatbarOn = True Then SetFormatBar

End Sub


Private Sub xxxmnuOpen_Click()
 If Text1.DataChanged And Len(Text1.Text) Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "iQ WordPad - Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 On Error GoTo Errhandler
 cmDlg.DialogTitle = "Open New Document"
 cmDlg.FileName = ""
 cmDlg.Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Word Document (*.doc)|*.doc|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Other Text (*.bas;*.bat;*.csv;*.dat;*.ini;*.lst)|*.bas;*.bat;*.csv;*.dat;*.ini;*.lst|All Files (*.*)|*.*"
 cmDlg.FilterIndex = 1
 cmDlg.ShowOpen
 
 Screen.MousePointer = 11
 
 IQFileName = cmDlg.FileName
 
 
 There = FileExists(IQFileName)
 If There Then
  Text1.Text = ""
   If LCase(Right$(IQFileName, 3)) = "rtf" Then
    Text1.LoadFile IQFileName, rtfRTF
   ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    picToolbar.Refresh
    frmConvert.Show
    DoEvents
    OpenWordDoc (IQFileName)
    frmConvert.Hide
   Else
    FileOpen
   End If
   
   ClearTabs
   Text1.SelBullet = False
   BulletNormal = False
   BulletNumber = False
   LineSpacing = 0
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
   CompactPath (LCase(IQFileName))
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
 Else
   MsgBox IQFileName & " Not Found!", vbOKOnly, "File Not Found"
 End If

Text1.SetFocus
picToolbar.Refresh
RulerBar.Refresh
picRuler.Refresh
picFormatbar.Refresh
If FormatbarOn = True Then SetFormatBar
DoEvents
Screen.MousePointer = 0

Exit Sub
 
Errhandler:
 If FormatbarOn = True Then SetFormatBar
 
 Screen.MousePointer = 0
 Exit Sub
 
End Sub

Private Sub xxxmnuPageSetup_Click()
 On Error GoTo Errhandler
 Dim prncheck As Variant
 
 prncheck = Printer.DeviceName

 frmPageSetup.Show 1
 
 EditMenuEnable
 Form_Resize
  
 Exit Sub
 
Errhandler:
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub
   
End Sub

Private Sub xxxmnuParagraph_Click()
 
 frmParagraph.Show 1
 
 EditMenuEnable
 If FormatbarOn = True Then SetFormatBar
 picToolbar.Refresh
 RulerBar.Refresh
 picRuler.Refresh
 picFormatbar.Refresh
End Sub

Private Sub xxxmnuPaste_Click()

'Could be anything in clipboard.  This will check
'and paste the correct type

 If Clipboard.GetFormat(vbCFRTF) Then
  Text1.SelRTF = Clipboard.GetText(vbCFRTF)
  Text1.SetFocus
  Exit Sub
 ElseIf Clipboard.GetFormat(vbCFBitmap) Then
  SendMessage Text1.hwnd, WM_PASTE, 0&, 0&
  Text1.SetFocus
  Exit Sub
 ElseIf Clipboard.GetFormat(vbCFText) And Not Clipboard.GetFormat(vbCFRTF) Then
  Text1.SelRTF = Clipboard.GetText(vbCFText)
  Text1.SetFocus
 End If

End Sub

Private Sub xxxmnuPrint_Click()
Dim prncheck As Variant

' Quick Print
 On Error GoTo Errhandler
 
 prncheck = Printer.DeviceName 'if no printer hooked up this check will get us out of here.

  If Len(Text1.SelText) > 1 Then
   
   ret = MsgBox("Do you wish to print selected text only?", vbYesNoCancel + vbQuestion, "iQ WordPad Print Text")
   
   If ret = vbCancel Then Exit Sub
   
   If ret = vbYes Then
    Call PrintSelected
    Exit Sub
   End If
   
   If ret = vbNo Then
    PrintRTF Text1, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
   End If
   
  Else

    PrintRTF Text1, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
    
  End If

  
  
 Exit Sub

Errhandler:
     
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub
   
End Sub

Private Sub xxxmnuPrinterSetup_Click()
 

'call print cmdlg
 On Error GoTo Errhandler
 
 cmDlg.flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDUseDevModeCopies
 cmDlg.ShowPrinter

 Printer.Copies = cmDlg.Copies
 
 For i = 1 To Printer.Copies
  Printer.Orientation = gOrientation
  Printer.PaperSize = gPaperSize
  PrintRTF Text1, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
 Next
 
 Exit Sub
 
Errhandler:
     
     If Err <> 32755 Then '32755 is cancel error
     
       MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
       
     End If
  Screen.MousePointer = 0
End Sub


Private Sub xxxmnuPrintPreview_Click()
 On Error GoTo Errhandler
 Dim prncheck As Variant
 
 prncheck = Printer.DeviceName
 
    gprint = False
    
    frmPrintPreview.Show 1
    
    If gprint Then
     xxxmnuPrint_Click
     gprint = False
    End If
 
 Exit Sub
 
Errhandler:
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub

End Sub

Private Sub xxxmnuRedo_Click()
End Sub

Private Sub xxxmnuReplace_Click()

 'Show Find/Replace API dialog
   Dim s As String
  If Text1.SelLength > 0 Then s = Text1.SelText Else s = ""
  ShowFind Me, Text1, 0, s, True, ""
  
End Sub

Private Sub xxxmnuRuler_Click()
 If RulerOn = False Then
  RulerOn = True
  xxxmnuRuler.Checked = True
  RulerBar.Visible = True
  Form_Resize
 Else
  RulerOn = False
  xxxmnuRuler.Checked = False
  RulerBar.Visible = False
  Form_Resize
 End If
End Sub

Private Sub xxxmnuRulerOption_Click(Index As Integer)
 
 xxxmnuRulerOption(0).Checked = False
 xxxmnuRulerOption(1).Checked = False
 
 picRuler.Picture = imgRuler(Index).Picture
 iQMeasurement = Index
 
 xxxmnuRulerOption(Index).Checked = True
  
End Sub

Private Sub xxxmnuSave_Click()
 On Error GoTo Errhandler
 
 If IQFileName = "Untitled" Then
  xxxmnuSaveAs_Click
  Exit Sub
 End If

 If LCase(Right$(IQFileName, 3)) = "rtf" Then
  Text1.SaveFile IQFileName, rtfRTF
 ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    picToolbar.Refresh
    frmConvert.Show
    SaveAsWordDoc (IQFileName)
    frmConvert.Hide
 Else
  Text1.SaveFile IQFileName, rtfText
 End If
 Text1.DataChanged = False
 Text1.SetFocus
 Exit Sub
 
Errhandler:

 If Err = 75 Then
  MsgBox "This document is Read-Only and cannot be modified.", vbOKOnly + vbCritical, "File not saved!"
  Exit Sub
 Else
  MsgBox "Error" & Str$(Err) & " - " & Error$, vbOKOnly + vbCritical, "File not saved!"
  Exit Sub
 End If
End Sub


Private Sub xxxmnuSaveAs_Click()
 On Error GoTo Errhandler
 fCancel = False
 Dim Temp As String
 Temp = IQFileName
 i = InStr(1, IQFileName, ".")
 If i Then Temp = Left$(IQFileName, i - 1)
 cmDlg.FileName = Temp
 cmDlg.flags = cdlOFNOverwritePrompt
 cmDlg.Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Word Document (*.doc)|*.doc|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Other Text (*.bas;*.bat;*.csv;*.dat;*.ini;*.lst)|*.bas;*.bat;*.csv;*.dat;*.ini;*.lst|All Files (*.*)|*.*"
 cmDlg.FilterIndex = 1
 cmDlg.ShowSave
 
 IQFileName = cmDlg.FileName

 If LCase(Right$(IQFileName, 3)) = "rtf" Then
  Text1.SaveFile IQFileName, rtfRTF
 ElseIf LCase(Right$(IQFileName, 3)) = "doc" Then
    frmConvert.Show
    picToolbar.Refresh
    DoEvents
    SaveAsWordDoc (IQFileName)
    frmConvert.Hide
 Else
  Text1.SaveFile IQFileName, rtfText
 End If
 
 Text1.DataChanged = False
 
 frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ WordPad"
 CompactPath (LCase(IQFileName))
 UpDateFileMenu IQFileName
 
 Exit Sub
 
Errhandler:
 If Err = 32755 Then fCancel = True
 If Err <> 32755 Then
  If Err = 75 Then
   MsgBox "This document is Read-Only and cannot be modified.", vbOKOnly + vbCritical, "File not saved!"
   Exit Sub
  Else
   MsgBox "Error" & Str$(Err) & " - " & Error$, vbOKOnly + vbCritical, "File not saved!"
   Exit Sub
  End If
 End If

 Exit Sub

End Sub

Private Sub xxxmnuSelectAll_Click()
    Dim ucharg As CharRange
    
    ucharg.cpMax = -1
    ucharg.cpMin = 0
    Call SendMessage(Text1.hwnd, EM_EXSETSEL, 0, ucharg)

End Sub

Private Sub xxxmnuSend_Click()
 Dim Start As Long
 
 If IQFileName = "Untitled" Then
  xxxmnuSaveAs_Click
  If fCancel = True Then Exit Sub
 Else
  xxxmnuSave_Click
 End If
 

   ' Start Outlook email.
   
    ShellExecute Me.hwnd, "Open", _
        "mailto:?subject=" & IQFileName & "&body=File Attached", _
        vbNullString, vbNullString, vbNormalFocus

    ' Wait until Outlook is ready.
    While ret = 0
        DoEvents
        ret = FindWindow(vbNullString, IQFileName)
    Wend
    
    Start = Timer + 0.3
    While Start > Timer
     DoEvents
    Wend
    
    ' Send keys Alt-I-A, the zip file name,
    ' two TABs, and Enter.
    SendKeys "%ia" & IQFileName & "{TAB}{TAB}{ENTER}", True
    
End Sub

Private Sub xxxmnuSpellCheck_Click()

'--------------------------------
'Code to use WSpell 3rd party control
'--------------------------------

'  On Error GoTo OpenError
'  Dim Scount As Long
  
  'Check to see if checking string or whole doc
'  If Text1.SelLength > 1 Then
    'checking selected text only
    'Dim text As String

    ' Get the contents of the text box into a string.

'    txtInsert.TextRTF = Text1.SelRTF
'    WSpell1.TextControlHWnd = txtInsert.hwnd

    ' Check the spelling of the string. Note that WSpell1's ShowContext
    ' property is set to True. This causes
    ' a context display to appear in the spell-check dialog box.
    'WSpell1.ShowContext = True
    'WSpell1.text = text
'   ret = WSpell1.Start
'    If (ret >= 0) Then
        ' The user didn't cancel, so put the correct string back into the text box.
'        Text1.SelText = txtInsert.TextRTF
'    End If
'   Else
    'check whole document
'    WSpell1.ShowContext = False
'    WSpell1.TextControlHWnd = Text1.hwnd
'    Call WSpell1.Start
'   End If
'    Scount = WSpell1.WordsReplacedCount
'    MyMsg = "Spell check completed." & vbCrLf
'    MyMsg = MyMsg & "Words replaced:" & Str$(Scount) & ".  "
    
'    MsgBox MyMsg, vbOKOnly + vbInformation, "iQ WordPad"
'    Exit Sub

'--------------------------------
'Code to use MS Word Spell Check
'--------------------------------
    Dim oWord As Object
    Dim oTmpDoc As Object
    Dim lOrigTop As Long
      
    On Error GoTo OpenError
    Screen.MousePointer = 11
    
    ' Create a Word document object...
    Set oWord = CreateObject("Word.Application")
    Set oTmpDoc = oWord.Documents.Add
    oWord.Visible = False
    
   ' Position Word off screen to avoid having document visible...
    oWord.WindowState = 0
    oWord.Top = -3000
    LockWindowUpdate Text1.hwnd
    Text1.HideSelection = True
    
        ' copy the contents of the text box to the clipboard
    If Text1.SelLength < 2 Then
     Text1.SelStart = 0
     Text1.SelLength = Len(Text1.TextRTF)
    End If
    
    'be sure to use RTF so we keep formatting
    Clipboard.Clear
    Clipboard.SetText Text1.SelRTF, vbCFRTF

    ' Assign the text to the document and check spelling...
    Screen.MousePointer = 0
    With oTmpDoc
        .Content.Paste
        .Activate
        .CheckSpelling
       
       ' .CheckGrammar 'uncomment this line to include grammarcheck
       
       ' After user has made changes, use the clipboard to
       ' transfer the contents back to the text box
        .Content.Copy
        
         Text1.SelRTF = Clipboard.GetText(vbCFRTF)

        ' Close the document and exit Word...
        .Saved = True
         Clipboard.Clear
        .Close
      End With
      
      'release the object
       Set oTmpDoc = Nothing
       oWord.Quit
       Set oWord = Nothing
      
      'Now tell the user we're done
       LockWindowUpdate 0
       SendKeys "{BACKSPACE}" 'get rid of extra return/linefeed added by word
       DoEvents
       MsgBox "Spell check is complete.", vbOKOnly + vbInformation, "iQ WordPad"
       Text1.HideSelection = False

      Exit Sub

'--------------------------------


'==========================================
'code for Spell Check Anywhere program
'==========================================
' If Len(Text1.TextRTF) < 2 Then Exit Sub
' Clipboard.Clear
 
' If Text1.SelLength < 2 Then
'   Text1.SelStart = 3
'   Text1.SetFocus
'   DoEvents
' End If
 
' SendKeys "{F11}", True
'
 'For i = 1 To 10000
' Next
' Exit sub
'==========================================


OpenError:
 MsgBox Error$ & " - " & Err, vbOKOnly, "Spell Check Error"
 Text1.HideSelection = False
 LockWindowUpdate 0
 Exit Sub

End Sub

Private Sub xxxmnuStatus_Click()
 If StatusbarOn = False Then
  StatusbarOn = True
  xxxmnuStatus.Checked = True
  StatusBar1.Visible = True
  Form_Resize
  If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
 Else
  StatusbarOn = False
  xxxmnuStatus.Checked = False
  StatusBar1.Visible = False
  Form_Resize
 End If
End Sub

Private Sub xxxmnuSubScript_Click()
On Error Resume Next
 Dim Temp As Integer
 Temp = Text1.SelCharOffset
 If IsNull(Temp) Then Temp = 0
 Text1.SelCharOffset = -55 + Temp
End Sub

Private Sub xxxmnuSuperScript_Click()
On Error Resume Next
 Dim Temp As Integer
 Temp = Text1.SelCharOffset
 If IsNull(Temp) Then Temp = 0
 Text1.SelCharOffset = 55 + Temp
 
End Sub

Private Sub xxxmnuTable_Click()
 frmTables.Show 1
 If fCancel = True Then Exit Sub 'they hit cancel button
 Dim j As Integer
 Dim xstart As Long
 xstart = Text1.SelStart
  
  With RTFtable
    'set the size of the table
    .Columns = tCol
    .Rows = tRow
    .isCentered = tCenter
    
    For i = 1 To tRow
     For j = 1 To tCol
      .Column(j).xWidth = tColWidth(j)
     Next j
    Next i
  LockWindowUpdate Text1.hwnd
    .InsertTable Text1
 
  End With

 Text1.SelStart = xstart + 2
 Text1.SetFocus
 LockWindowUpdate 0
  
End Sub

Private Sub xxxmnuTabs_Click()
 frmTabs.Show 1
 ShowTabs
 
End Sub

Private Sub xxxmnuTextProperties_Click()
 'show property dialog
 ret = ShowFileProp(IQFileName, frmMainText)
 
End Sub

Private Sub xxxmnuTimeDate_Click()
 'Text1.SelText = Format$(Now, "h:mm AMPM m/dd/yyyy")
 frmDateTime.Show 1
End Sub

Private Sub xxxmnuTool_Click()
 If ToolbarOn = False Then
  ToolbarOn = True
  xxxmnuTool.Checked = True
  picToolbar.Visible = True
  Form_Resize
  picToolbar.Refresh
  
 Else
  ToolbarOn = False
  xxxmnuTool.Checked = False
  picToolbar.Visible = False
  Form_Resize
 End If
End Sub

Private Sub xxxmnuUndo_Click()

 SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
 EditMenuEnable
End Sub

Private Sub xxxmnuViewClipboard_Click()
 Shell "clipbrd.exe", vbNormalFocus
End Sub

Private Sub xxxmnuWordCount_Click()
On Error Resume Next
 Dim WCount As Long
 Dim LnCount As Long
 Dim CharCount As Long
 CharCount = 0
 WCount = WordCount(Text1.Text)
 LnCount = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0, 0&)
 CharCount = SendMessageLong(Text1.hwnd, WM_GETTEXTLENGTH, 0, 0)
 CharCount = Format(CharCount, "###,###,###,###,###")
 MyMsg = ""
 MyMsg = IQFileName & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Word Count:" & Str$(WCount) & "   " & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Line Count:" & Str$(LnCount) & "   " & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Character Count:" & Str$(CharCount) & "   " & vbCrLf & vbCrLf
 
 MsgBox MyMsg, vbOKOnly + vbInformation, "iQ WordPad Statistics"

End Sub

Private Sub xxxmnuWordwrap_Click(Index As Integer)
 For i = 0 To 2
  xxxmnuWordwrap(i).Checked = False
 Next
 
 xxxmnuWordwrap(Index).Checked = True
 WrapOn = Index
    
 Select Case Index
  
  Case 0 'Wrap to Window
    Text1.RightMargin = 0

  Case 1 'Wrap to Ruler
    Text1.RightMargin = gMargin + 90
    
  Case 2 'no wrap
    Text1.RightMargin = 200000
 
 End Select
 
 Form_Resize
 
End Sub


