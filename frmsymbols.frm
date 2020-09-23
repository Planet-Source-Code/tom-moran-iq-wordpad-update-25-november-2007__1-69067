VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Symbols and Characters"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   300
      Width           =   1275
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4980
      TabIndex        =   6
      Top             =   300
      Width           =   1215
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   24
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Click on desired symbol.  Click Insert or double click symbol to insert in text.  Click on Close button to exit."
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   5040
      Width           =   7380
   End
   Begin VB.Label lblKeyStroke 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alt+0174 "
      Height          =   255
      Left            =   3540
      TabIndex        =   3
      Top             =   360
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keystroke:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   60
      Width           =   990
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboFont_Click()
 Grid1.Font.Name = cboFont.Text
 Grid1.Refresh
 Grid1.SetFocus
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdInsert_Click()
 frmMainText.Text1.SelFontName = cboFont.Text
 frmMainText.Text1.SelText = Grid1.Text
End Sub

Private Sub Form_Load()
 Me.Top = frmMainText.Top + 300
 Me.Left = frmMainText.Left + 1800
 

 Dim Gridcount As Integer
 Dim i%, j%
    
    lblKeyStroke.Caption = ""
    For i = 1 To Screen.FontCount - 1
        If Screen.Fonts(i) <> "" Then cboFont.AddItem Screen.Fonts(i)
    Next i
    
    cboFont.Text = frmMainText.Text1.SelFontName
    Grid1.Font.Name = frmMainText.Text1.SelFontName
    Gridcount = 31
 
   For i = 0 To 22 ' Set row numbers.
 
    For j = 0 To 9 ' col numbers
     Gridcount = Gridcount + 1
     If Gridcount = 256 Then Exit For
     Grid1.Row = i
     Grid1.Col = j
     Grid1.ColWidth(j) = 720
     Grid1.ColAlignment(j) = 4
     Grid1.Text = Chr$(Gridcount)
    Next j
   
   Next i
   Screen.MousePointer = 0

End Sub

Private Sub Grid1_Click()
  Dim Char$
  Dim charnum As Integer
  
  Char$ = Grid1.Text
  If Char$ = "" Then Exit Sub

  charnum = Asc(Char$)
  If charnum > 126 Then
   lblKeyStroke.Caption = " Alt+0" + LTrim(Str$(charnum)) & " "
  Else
   lblKeyStroke.Caption = " Alt+" + LTrim(Str$(charnum)) & " "
  End If

End Sub

Private Sub Grid1_DblClick()
 frmMainText.Text1.SelFontName = cboFont.Text
 frmMainText.Text1.SelText = Grid1.Text
End Sub


