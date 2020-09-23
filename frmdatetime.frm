VERSION 5.00
Begin VB.Form frmDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Date and Time"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   2715
   End
   Begin iQWordPad.CandyButton cmdOkay 
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   660
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   3060
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Available Formats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOkay_Click()
 frmMainText.Text1.SelText = List1.List(List1.ListIndex) & " "
 Unload Me
End Sub

Private Sub Form_Load()
 Dim th As Variant
 Dim ext As Variant
 Dim A$
 
  th = Format(Now, "d")
 Select Case th
  Case 1, 21, 31
   ext = th + "st"
  Case 2, 22
   ext = th + "nd"
  Case 3, 23
   ext = th + "rd"
  Case Else
   ext = th + "th"
 End Select
 
 List1.AddItem Format(Now, "m/d/yy")

 List1.AddItem Format(Now, "mm/dd/yy")

 List1.AddItem Format(Now, "mm/dd/yyyy")

 List1.AddItem Format(Now, "mmmm dd, yyyy")
 
 A$ = Format(Now, "mmmm d, yyyy")
 pos = InStr(A$, " ")
 A$ = Left$(A$, pos) + ext + Right$(A$, 6)
 List1.AddItem A$

 List1.AddItem Format(Now, "ddd") + " - " + Format(Now, "mm/dd/yy")

 List1.AddItem Format(Now, "dddddd")

 List1.AddItem Now

 List1.AddItem Format(Now, "hh:mm am/pm")

 List1.AddItem Format(Now, "hh:mm AM/PM")

 List1.AddItem Format(Now, "h:mm:ss am/pm")

 List1.AddItem Format(Now, "hh:mm:ss AM/PM")

 List1.AddItem Format(Now, "Short Time")

 List1.AddItem Format(Now, "hh:mm:ss")

 List1.ListIndex = 0

End Sub
