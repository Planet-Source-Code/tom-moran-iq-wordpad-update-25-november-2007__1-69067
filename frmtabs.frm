VERSION 5.00
Begin VB.Form frmTabs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabs"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   Icon            =   "frmTabs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1620
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   300
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tab stop positions"
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin iQWordPad.CandyButton cmdSet 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Set"
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
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   2355
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2355
      End
      Begin iQWordPad.CandyButton cmdClear 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   2160
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clear"
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
   End
End
Attribute VB_Name = "frmTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
 Unload Me
End Sub



Private Sub cmdClear_Click()
 List1.RemoveItem List1.ListIndex
 Text1.SetFocus
End Sub

Private Sub cmdClearAll_Click()
    List1.Clear ' clear them
    cmdClearAll.Enabled = False
    Text1.Text = "" 'clear the tab text
    Text1.SetFocus 'set focus to the new tab box
End Sub

Private Sub cmdOk_Click()
 Dim pos As Integer
 Dim TValue As Double
 
 frmMainText.Text1.SelTabCount = List1.ListCount
 
 If iQMeasurement = 0 Then
   For i = 0 To List1.ListCount - 1
    pos = InStr(List1.List(i), " ")
    TValue = Val(Left$(List1.List(i), pos))
    frmMainText.Text1.SelTabs(i) = TValue * 1440
   Next
 Else
   For i = 0 To List1.ListCount - 1
    pos = InStr(List1.List(i), " ")
    TValue = Val(Left$(List1.List(i), pos))
    frmMainText.Text1.SelTabs(i) = (TValue * 1440) / 2.54
   Next
 End If
 
 Unload Me
 
 
End Sub

Private Sub cmdSet_Click()
 'make sure we are in range
  If Val(Text1.Text) < 0.1 Or Val(Text1.Text) > 22 Then
   MsgBox "Measurement must be between 0.1 and 22.", vbOKOnly + vbInformation, "Tab set error!"
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.Text)
   Text1.SetFocus
   Exit Sub
  End If
  If iQMeasurement = 0 Then
   Text1.Text = Text1.Text & " """
  Else
   Text1.Text = Text1.Text & " cm"
  End If
  
  If Len(Text1.Text) < 8 And iQMeasurement = 1 Then Text1.Text = "0" & Text1.Text
    
    List1.AddItem Text1.Text
    cmdClearAll.Enabled = True
    Text1.Text = ""
    Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Temp As String

If IsNull(frmMainText.Text1.SelTabCount) Then Exit Sub 'frmMainText.Text1.SelTabCount = 0

If frmMainText.Text1.SelTabCount > 0 Then
 cmdClearAll.Enabled = True
End If

 
    For i = 0 To frmMainText.Text1.SelTabCount - 1
      Dim sglTabValue As Single
      If iQMeasurement = 0 Then
       sglTabValue = frmMainText.Text1.SelTabs(i) / 1440#
       sglTabValue = CInt(sglTabValue * 100) / 100#
        If sglTabValue < 1 Then
         Temp = "0" & Trim(Str(sglTabValue)) & " """
        Else
         Temp = Trim(Str(sglTabValue)) & " """
        End If
      Else
       sglTabValue = (frmMainText.Text1.SelTabs(i) / 1440#) * 2.54
       sglTabValue = CInt(sglTabValue * 100) / 100#
        If sglTabValue < 1 Then
         Temp = "0" & Trim(Str(sglTabValue)) & " cm"
        Else
         Temp = Trim(Str(sglTabValue)) & " cm"
        End If
        If Len(Temp) < 8 Then Temp = "0" & Temp
      End If
      List1.AddItem Temp
    Next i
    
End Sub

Private Sub List1_GotFocus()
 If List1.ListCount > 0 Then
  cmdClear.Enabled = True
 Else
  cmdClear.Enabled = False
 End If
 
 
End Sub


Private Sub Text1_Change()
 If Len(Text1) Then
  cmdSet.Enabled = True
 Else
  cmdSet.Enabled = False
 End If
End Sub

Private Sub Text1_GotFocus()
 cmdClear.Enabled = False
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  cmdSet_Click
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 45 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub
