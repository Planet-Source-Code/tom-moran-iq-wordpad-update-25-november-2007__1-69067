VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goto line"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin iQWordPad.CandyButton cmdOkay 
      Height          =   375
      Left            =   780
      TabIndex        =   2
      Top             =   900
      Width           =   1275
      _ExtentX        =   2249
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
      Caption         =   "Okay"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtGoto 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
   Begin iQWordPad.CandyButton cmdCancel 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   900
      Width           =   1275
      _ExtentX        =   2249
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line Number:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOkay_Click()
 Dim gotoline As Long
            'Get pos of start of the line
            gotoline = SendMessage(frmMainText.Text1.hWnd, EM_LINEINDEX, txtGoto.Text - 1, 0&)
            If gotoline = -1 Then 'Invalid line number
                MsgBox "Line number out of range", 0, "iQ TextVue - Goto Line"
                Exit Sub
            End If
            frmMainText.Text1.SelStart = gotoline 'Go To line
            DoEvents
            Unload Me
End Sub

Private Sub Form_Load()
 Me.Left = frmMainText.Left + 10
 Me.Top = frmMainText.Top + frmMainText.Text1.Top
     
    Dim CurrentLine As Long
    CurrentLine = SendMessage(frmMainText.Text1.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    txtGoto.Text = Trim(Str$(CurrentLine))
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto.Text)
    
    
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  cmdOkay.SetFocus
  cmdOkay_Click
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If

 
End Sub

