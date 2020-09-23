VERSION 5.00
Begin VB.Form frmColorPick 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Color Selection Tool"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmColorPick.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   2205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin iQWordPad.CandyButton cmdCustom 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2100
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      More..."
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   16711680
      Picture         =   "frmColorPick.frx":268CA
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin iQWordPad.CandyButton cmdCancel 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   15
      Left            =   1680
      TabIndex        =   17
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   14
      Left            =   1200
      TabIndex        =   16
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   13
      Left            =   720
      TabIndex        =   15
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   12
      Left            =   240
      TabIndex        =   14
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   1680
      TabIndex        =   13
      Top             =   1140
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   10
      Left            =   1200
      TabIndex        =   12
      Top             =   1140
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   720
      TabIndex        =   11
      Top             =   1140
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   1140
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   1680
      TabIndex        =   9
      Top             =   660
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   1200
      TabIndex        =   8
      Top             =   660
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   720
      TabIndex        =   7
      Top             =   660
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   180
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   180
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   180
      Width           =   315
   End
   Begin VB.Label shpColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   315
   End
End
Attribute VB_Name = "frmColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdCustom_Click()
On Error GoTo Errhandler

 frmMainText.cmDlg.flags = cdlCCFullOpen
 frmMainText.cmDlg.ShowColor
 
 If HighColor = False Then
   frmMainText.Text1.SelColor = frmMainText.cmDlg.Color
 Else
   ColorHighLite = frmMainText.cmDlg.Color
   HighliteText (ColorHighLite)
 End If
 
 Unload Me
 Exit Sub

Errhandler:
 
 Exit Sub

End Sub

Private Sub Form_Load()
 
Call GetCursorPos(Pnt)
Me.Left = Pnt.mx * 15
Me.Top = Pnt.my * 15
 
 For i = 0 To 15
  shpColor(i).BackColor = QBColor(i)
 Next
  
End Sub

Private Sub Label1_Click()

End Sub

Private Sub shpColor_Click(Index As Integer)

 If HighColor = False Then
   frmMainText.Text1.SelColor = QBColor(Index)
 Else
   ColorHighLite = QBColor(Index)
   HighliteText (ColorHighLite)
 End If
 
 Unload Me
 Exit Sub
End Sub


