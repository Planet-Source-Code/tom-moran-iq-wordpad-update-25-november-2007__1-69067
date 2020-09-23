VERSION 5.00
Begin VB.Form frmParagraph 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paragraph"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "frmParagraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLineSpacing 
      Height          =   315
      Left            =   1260
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1635
   End
   Begin VB.ComboBox cboAlignment 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2100
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indentation in Inches"
      ForeColor       =   &H8000000D&
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2835
      Begin VB.TextBox txtHanging 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "txtHanging"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox txtRight 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Text            =   "txtRight"
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtLeft 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Text            =   "txtLeft"
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Line:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   315
      End
   End
   Begin iQWordPad.CandyButton cmdOkay 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
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
   Begin iQWordPad.CandyButton cmdCancel 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1140
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Line Spacing:"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment:"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frmParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
 Unload Me
 
End Sub

Private Sub cmdOkay_Click()
Dim LeftIndent As Long
Dim HangIndent As Long

If iQMeasurement = 0 Then
  LeftIndent = Val(txtLeft.Text) * 1440
  HangIndent = Val(txtHanging.Text) * -1440
  frmMainText.Text1.SelRightIndent = Val(txtRight.Text) * 1440
Else
  LeftIndent = (Val(txtLeft.Text) * 1440) / 2.54
  HangIndent = (Val(txtHanging.Text) * -1440) / 2.54
  frmMainText.Text1.SelRightIndent = (Val(txtRight.Text) * 1440) / 2.54
End If
 
 frmMainText.Text1.SelIndent = (LeftIndent - HangIndent)
 frmMainText.Text1.SelHangingIndent = HangIndent
 frmMainText.Text1.SelAlignment = cboAlignment.ListIndex
 
 LineSpacing = cboLineSpacing.ListIndex
 SelLineSpacing frmMainText.Text1, LineSpacing
 
 Unload Me
 
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Numb As Variant
Dim HNumb As Variant
Dim pos As Integer

If iQMeasurement = 0 Then
 Frame1.Caption = "Indent in Inches"
 HNumb = frmMainText.Text1.SelHangingIndent / 1440
Else
 Frame1.Caption = "Indent in Centimeters"
 HNumb = (frmMainText.Text1.SelHangingIndent / 1440) * 2.54
End If
If IsNull(HNumb) Then HNumb = 0
pos = InStr(HNumb, ".")
If pos Then
 HNumb = Left$(HNumb, pos + 2)
End If
txtHanging.Text = Abs(HNumb) '+ 0.02

If iQMeasurement = 0 Then
 Numb = frmMainText.Text1.SelIndent / 1440
Else
 Numb = (frmMainText.Text1.SelIndent / 1440) * 2.54
End If
If IsNull(Numb) Then Numb = 0
pos = InStr(Numb, ".")
If pos Then
 Numb = Left$(Numb, pos + 2)
End If
 Numb = Val(Numb) + Val(HNumb)
txtLeft.Text = Val(Numb) '- 0.02

If iQMeasurement = 0 Then
 Numb = frmMainText.Text1.SelRightIndent / 1440
Else
 Numb = (frmMainText.Text1.SelRightIndent / 1440) * 2.54
End If

If IsNull(Numb) Then Numb = 0
pos = InStr(Numb, ".")
If pos Then
 Numb = Left$(Numb, pos + 2)
End If
txtRight.Text = Numb

'Get Alignment
 cboAlignment.AddItem "Left"
 cboAlignment.AddItem "Right"
 cboAlignment.AddItem "Center"
 
 If IsNull(frmMainText.Text1.SelAlignment) Then
  cboAlignment.ListIndex = 0
 Else
  cboAlignment.ListIndex = frmMainText.Text1.SelAlignment
 End If
 
'Get LineSpacing
 cboLineSpacing.AddItem "Single Space"
 cboLineSpacing.AddItem "1.5 line Space"
 cboLineSpacing.AddItem "Double Space"
 
 If IsNull(LineSpacing) Then
  cboLineSpacing.ListIndex = 0
 Else
  cboLineSpacing.ListIndex = LineSpacing
 End If
 
   
End Sub


Private Sub txtHanging_GotFocus()
 txtHanging.SelStart = 0
 txtHanging.SelLength = Len(txtHanging.Text)
End Sub

Private Sub txtHanging_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  cmdOkay.SetFocus
  cmdOkay_Click
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtHanging_LostFocus()
 If Val(txtHanging.Text) < -22 Or Val(txtHanging.Text) > 22 Then
  MsgBox "Number must be between -22 and 22", vbOKOnly + vbCritical, "Paragraph Error"
  txtHanging.Text = "0"
  txtHanging.SetFocus
 End If
End Sub

Private Sub txtLeft_GotFocus()
 txtLeft.SelStart = 0
 txtLeft.SelLength = Len(txtLeft.Text)
 
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtRight.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub

Private Sub txtLeft_LostFocus()
 If Val(txtLeft.Text) < -22 Or Val(txtLeft.Text) > 22 Then
  MsgBox "Number must be between -22 and 22", vbOKOnly + vbCritical, "Paragraph Error"
  txtLeft.Text = "0"
  txtLeft.SetFocus
 End If
End Sub

Private Sub txtRight_GotFocus()
 txtRight.SelStart = 0
 txtRight.SelLength = Len(txtRight.Text)
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtHanging.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtRight_LostFocus()
 If Val(txtRight.Text) < -22 Or Val(txtRight.Text) > 22 Then
  MsgBox "Number must be between -22 and 22", vbOKOnly + vbCritical, "Paragraph Error"
  txtRight.Text = "0"
  txtRight.SetFocus
 End If
End Sub


