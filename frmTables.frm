VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Table"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   Icon            =   "frmTables.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCenter 
      Caption         =   "Center table"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2340
      Width           =   2235
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3060
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   435
      Left            =   3060
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Table Dimensions (inches)"
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   2715
      Begin iQWordPad.CandyButton cmdSetWidth 
         Height          =   255
         Left            =   2220
         TabIndex        =   10
         Top             =   1380
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.TextBox txtWidth 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "txtWidth"
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txtCol 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Text            =   "txtCol"
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtRows 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Text            =   "txtRows"
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Col Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Columns:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rows:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tColWid(1 To 20) As Double

Private Sub cmdCancel_Click()
fCancel = True
Unload Me
End Sub

Private Sub cmdOkay_Click()
 tRow = Val(txtRows.Text)
 tCol = Val(txtCol.Text)
 For i = 1 To Val(txtCol)
  If tColWidth(i) = 0 Then tColWidth(i) = Val(txtWidth) * 1440
 Next
 If chkCenter.Value = 1 Then tCenter = True Else tCenter = False
 fCancel = False
 Unload Me
 
End Sub

Private Sub cmdSetWidth_Click()
Dim Temp As String
Dim TMsg As String

TMsg = "Enter the width in inches for column"
For i = 1 To Val(txtCol)
 Temp = InputBox(TMsg & Str$(i), "Set width for column" & Str$(i), Val(txtWidth))
 If Val(Temp) < 0.25 Or Val(Temp) > 12 Then
  MsgBox "Invalid measurement!  Enter measurement between .25 and 12.", vbOKOnly + vbCritical, "iQ WordPad Table Error"
  Exit Sub
 End If
 tColWidth(i) = Val(Temp) * 1440
 chkCenter.SetFocus
Next

 
End Sub

Private Sub Form_Load()
 fCancel = True 'in case they click x to exit
 
 'Defaults
 txtRows = 2
 txtCol = 2
 txtWidth = 1
 tColWidth(1) = 1 * 1440
 tColWidth(2) = 1 * 1440
 For i = 3 To 20
  tColWidth(i) = 0
 Next
 
 
End Sub

Private Sub txtCol_GotFocus()
 txtCol.SelStart = 0
 txtCol.SelLength = Len(txtCol.Text)
End Sub


Private Sub txtCol_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtWidth.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtCol_LostFocus()
 If Val(txtCol.Text) < 1 Or Val(txtCol.Text) > 22 Then
  MsgBox "Number must be between 1 and 22", vbOKOnly + vbCritical, "Table Error"
  txtCol.Text = "1"
  txtCol.SetFocus
 End If
End Sub


Private Sub txtRows_GotFocus()
 txtRows.SelStart = 0
 txtRows.SelLength = Len(txtRows.Text)
End Sub


Private Sub txtRows_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtCol.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtRows_LostFocus()
 If Val(txtRows.Text) < 1 Or Val(txtRows.Text) > 22 Then
  MsgBox "Number must be between 1 and 22", vbOKOnly + vbCritical, "Table Error"
  txtRows.Text = "1"
  txtRows.SetFocus
 End If
End Sub

Private Sub txtWidth_GotFocus()
 txtWidth.SelStart = 0
 txtWidth.SelLength = Len(txtWidth.Text)
End Sub


Private Sub txtWidth_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  cmdOkay.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub

Private Sub txtWidth_LostFocus()
 If Val(txtWidth.Text) < 0.25 Or Val(txtWidth.Text) > 12 Then
  MsgBox "Number must be between .25 and 12", vbOKOnly + vbCritical, "Table Error"
  txtWidth.Text = "1"
  txtWidth.SetFocus
  Exit Sub
 End If

 For i = 1 To Val(txtCol)
  tColWidth(i) = Val(txtWidth) * 1440
 Next
End Sub


