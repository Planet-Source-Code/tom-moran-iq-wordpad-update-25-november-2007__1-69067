VERSION 5.00
Begin VB.Form frmPrnSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Default Printer"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmPrnSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   2595
   End
   Begin iQWordPad.CandyButton cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2220
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
      Caption         =   "Close"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin iQWordPad.CandyButton cmdOkay 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1380
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
      Caption         =   "Set"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Printers Available"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblPrintername 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1620
      TabIndex        =   2
      Top             =   180
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Selected:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1440
   End
End
Attribute VB_Name = "frmPrnSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrnError As Long



Private Sub cmdCancel_Click()
 Unload Me
End Sub


Private Sub cmdOkay_Click()
On Error GoTo Errhandler
 
  i = List1.ListIndex
  
  Set Printer = Printers(i)
  
  lblPrintername.Caption = Printer.DeviceName
  
  Unload Me
  
  Exit Sub
  
Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 Unload Me
 
End Sub

Private Sub Form_Activate()

If PrnError Then Unload Me


 
End Sub

Private Sub Form_Load()
On Error GoTo Errhandler

 lblPrintername.Caption = Printer.DeviceName

    Dim p As Printer
    
    For Each p In Printers
     List1.AddItem p.DeviceName
    Next
 For i = 0 To List1.ListCount - 1
  If List1.List(i) = Printer.DeviceName Then
   List1.ListIndex = i
   Exit Sub
  End If
 Next

Exit Sub
 
Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 PrnError = True
 
End Sub

Private Sub List1_DblClick()
On Error Resume Next
 
  i = List1.ListIndex
  
  Set Printer = Printers(i)
  
  lblPrintername.Caption = Printer.DeviceName
  
  Unload Me

 End Sub

