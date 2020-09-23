VERSION 5.00
Begin VB.Form frmPrintPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   7500
   ClientLeft      =   1125
   ClientTop       =   1500
   ClientWidth     =   10500
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   2  'CenterScreen
   Begin iQWordPad.CandyButton cmdZoomOut 
      Height          =   315
      Left            =   420
      TabIndex        =   18
      Top             =   60
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmPrintPreview.frx":058A
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
   Begin iQWordPad.CandyButton cmdClose 
      Height          =   315
      Left            =   8940
      TabIndex        =   14
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Close"
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
   Begin VB.ComboBox cboScale 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   60
      Width           =   855
   End
   Begin VB.ComboBox cboPageNo 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   60
      Width           =   825
   End
   Begin VB.PictureBox PicZ 
      BackColor       =   &H009E9E9E&
      Height          =   6705
      Left            =   60
      ScaleHeight     =   6645
      ScaleWidth      =   10065
      TabIndex        =   2
      Top             =   480
      Width           =   10125
      Begin VB.PictureBox Pic5 
         BackColor       =   &H80000009&
         Height          =   2295
         Left            =   300
         ScaleHeight     =   2235
         ScaleWidth      =   2595
         TabIndex        =   9
         Top             =   120
         Width           =   2655
      End
      Begin VB.PictureBox Pic4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   2715
         Left            =   270
         ScaleHeight     =   2655
         ScaleWidth      =   3015
         TabIndex        =   8
         Top             =   180
         Width           =   3075
      End
      Begin VB.PictureBox Pic3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   3285
         Left            =   240
         ScaleHeight     =   3225
         ScaleWidth      =   3765
         TabIndex        =   7
         Top             =   150
         Width           =   3825
      End
      Begin VB.PictureBox Pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   3795
         Left            =   240
         ScaleHeight     =   3735
         ScaleWidth      =   4515
         TabIndex        =   6
         Top             =   60
         Width           =   4575
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   4215
         Left            =   180
         ScaleHeight     =   4155
         ScaleWidth      =   5325
         TabIndex        =   5
         Top             =   60
         Width           =   5385
      End
      Begin VB.PictureBox PicX 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   4695
         Left            =   150
         ScaleHeight     =   4635
         ScaleWidth      =   6015
         TabIndex        =   4
         Top             =   60
         Width           =   6075
      End
      Begin VB.PictureBox picP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   6090
         Left            =   120
         ScaleHeight     =   6030
         ScaleWidth      =   6885
         TabIndex        =   3
         Top             =   60
         Width           =   6945
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6615
      Left            =   10200
      Max             =   500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   60
      Max             =   500
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7200
      Width           =   10125
   End
   Begin iQWordPad.CandyButton cmdPrint 
      Height          =   315
      Left            =   7440
      TabIndex        =   15
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "    Print"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmPrintPreview.frx":0B24
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
   Begin iQWordPad.CandyButton cmdNextPage 
      Height          =   315
      Left            =   4560
      TabIndex        =   16
      Top             =   60
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  Next >"
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
   Begin iQWordPad.CandyButton cmdPrevPage 
      Height          =   315
      Left            =   3360
      TabIndex        =   17
      Top             =   60
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "< Previous"
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
   Begin iQWordPad.CandyButton cmdZoomIn 
      Height          =   315
      Left            =   1080
      TabIndex        =   19
      Top             =   60
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmPrintPreview.frx":0EBE
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
   Begin VB.Image Image1 
      Height          =   465
      Left            =   0
      Picture         =   "frmPrintPreview.frx":1458
      Top             =   0
      Width           =   11355
   End
   Begin VB.Label lblPercent 
      Caption         =   "%"
      Height          =   225
      Left            =   2220
      TabIndex        =   13
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblTotalPages 
      Caption         =   "of 4000 pages"
      Height          =   225
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Originally was RTPreview.frm  (Also known as DocPreview)  By Herman Liu
'   Modified to conform to Microsofts WYSIWIG code.  Redesigned, changed zoom
'   percentages, made for SDI instead of MDE and centering of preview pages.  Update interface

Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    firstChar As Long         ' First character of range (0 for start of doc)
    lastChar As Long          ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long               ' Actual DC to draw on
    hdcTarget As Long         ' Target DC for determining text formatting
    rectRegion As Rect        ' Region of the DC to draw to (in twips)
    rectPage As Rect          ' Page size of the entire DC (in twips)
    mCharRange As CharRange   ' Range of text to draw (see above user type)
End Type

'-------------------------------------------------------------------------------------------
' By using message in VB, it is possible to make a RTB support WYSIWYG display and output:
' EM_SETTARGETDEVICE tells a RTB to base its display on a target device.
' EM_FORMATRANGE sends a page at a time to an output device at specified coordinates.
'-------------------------------------------------------------------------------------------

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const WM_PASTE = &H302

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, Ip As Any) As Long
     
Dim mFormatRange As FormatRange
Dim rectDrawTo As Rect
Dim rectPage As Rect
Dim TextLength As Long
Dim newStartPos As Long
Dim dumpaway As Long
     
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
     (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
     ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
'-------------------------------------------------------------------------------------------------------------------

Dim mSizeNo As Integer
Dim currTotalPages As Integer

   Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Me.Caption = " " & LastPart(IQFileName) & " - iQ WordPad:  Print Preview"
   gprint = False
     ' we don't want the sizes to change after they have been appropriately sized
   picP.AutoSize = False             ' For print intermediary, always invisible
   PicX.AutoSize = False             ' For diaplay intermediary, always invisible
   Pic1.AutoSize = False             ' As 100%
   Pic2.AutoSize = False             ' As 75%
   Pic3.AutoSize = False             ' As 50%
   Pic4.AutoSize = False             ' As 40%
   Pic5.AutoSize = False             ' As 25%
   
      ' By default VB prints in twips. If a Picturebox is using pixels, we have to
      ' convert twips to pixels.  Therefore we fix the size of Pictureboxes before
      ' setting its ScaleMode to pixel.
      
   Dim mNormalWidth, mNormalHeight
   Dim mAdjFactor
   Dim mRect, mNewRect, mfactor
   Dim mpage As Integer
   
      ' Render document size in line with that of the printer (but note that doc is
      ' shown on screen without print margins)
   DocWYSIWYG frmMainText.Text1
   
      ' Obtain size of the printer/paper
   mNormalWidth = Printer.ScaleWidth
   mNormalHeight = Printer.ScaleHeight
   
      ' Mark down rectangle area, see remarks later
   mRect = mNormalWidth * mNormalHeight
   
      ' Make the invisible PicX of the same size as printer
   PicX.Width = PicX.Width - PicX.ScaleWidth + mNormalWidth
   PicX.Height = PicX.Height - PicX.ScaleHeight + mNormalHeight
   
     ' Percentage may be expressed in terms of original area (in that case, we have to
     ' derive the width and height from the computed area), or in terms of width and
     ' height themselves.  Here, to stress the point, we apply the percentage in terms
     ' of the area for sizes over 100%, but apply the percentage in terms of the width
     ' width and height themselves for sizes which are below 100%.
   
   'mNewRect = mRect * (150 / 100)
  ' mfactor = Sqr(mNewRect / (mNormalWidth * mNormalHeight))
  ' Pic1.Width = CInt(mNormalWidth * mfactor)
  ' Pic1.Height = CInt(mNormalHeight * mfactor)
      
   Pic1.Width = PicX.Width
   Pic1.Height = PicX.Height
          
   Pic2.Width = CInt(mNormalWidth * 75 / 100)
   Pic2.Height = CInt(mNormalHeight * 75 / 100)
   
   Pic3.Width = CInt(mNormalWidth * 50 / 100)
   Pic3.Height = CInt(mNormalHeight * 50 / 100)
   
   Pic4.Width = CInt(mNormalWidth * 40 / 100)
   Pic4.Height = CInt(mNormalHeight * 40 / 100)
   
   Pic5.Width = CInt(mNormalWidth * 25 / 100)
   Pic5.Height = CInt(mNormalHeight * 25 / 100)
   
      ' Set ScaleMode to pixels.
   frmPrintPreview.ScaleMode = vbPixels
   PicZ.ScaleMode = vbPixels
   PicX.ScaleMode = vbPixels
   Pic1.ScaleMode = vbPixels
   Pic2.ScaleMode = vbPixels
   Pic3.ScaleMode = vbPixels
   Pic4.ScaleMode = vbPixels
   Pic5.ScaleMode = vbPixels
   
     ' Set AutoRedraw to True
   PicZ.AutoRedraw = True
   picP.AutoRedraw = True
   PicX.AutoRedraw = True
   Pic1.AutoRedraw = True
   Pic2.AutoRedraw = True
   Pic3.AutoRedraw = True
   Pic4.AutoRedraw = True
   Pic5.AutoRedraw = True
   
    ' Before showing first page, test how many pages are there in total in RTB.
   currTotalPages = PageCtnProc(frmPrintPreview.PicX)
    ' Display the No. of total pages available
   lblTotalPages.Caption = "of " & CStr(currTotalPages) & " pages"
    ' Enable/disable page movement buttons
   setPageButtons
   
   Dim i As Integer
   cboPageNo.Clear
   For i = 1 To currTotalPages
       cboPageNo.AddItem i
   Next i
   cboPageNo.Text = cboPageNo.List(0)
   
      ' For ComboBox list
   cboScale.AddItem "100"
   cboScale.AddItem "75"
   cboScale.AddItem "50"
   cboScale.AddItem "40"
   cboScale.AddItem "25"
   cboScale.Text = cboScale.List(3)      ' i.e. 25%
    
      ' Print to PicX first, then project to other pictureboxes according to the sizes
      ' they play.
   mpage = 1
   PrintPreviewPage frmPrintPreview.PicX, mpage
    
      ' Now blast to wanted sizes.
    For i = 1 To 5
        DoEvents
        If MakeSizes(i) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Next
    Screen.MousePointer = vbDefault
     
     ' Start display of preview screen.
    PicZ.Visible = True
    picP.Visible = False
    PicX.Visible = False
    
    mSizeNo = 4             ' i.e. cboScale.List=3, 40%
    ChangePreview
End Sub



Private Sub cboPageNo_click()
    Dim mpage As Integer
    mpage = cboPageNo.ListIndex + 1
    setPageButtons
    Screen.MousePointer = vbHourglass
     ' Print a new page to PicX
    PrintPreviewPage frmPrintPreview.PicX, mpage
     ' Blast to various sizes.
    Dim i
    For i = 1 To 5
        DoEvents
        If MakeSizes(i) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Next
    ChangePreview
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdPrevPage_Click()
    If currTotalPages = 1 Then
        Exit Sub
    Else
        If Val(cboPageNo.Text) > 1 Then
            cboPageNo.Text = cboPageNo.List(cboPageNo.ListIndex - 1)
        End If
    End If
End Sub



Private Sub cmdNextPage_Click()
    If currTotalPages = 1 Then
        Exit Sub
    Else
        If Val(cboPageNo.Text) < currTotalPages Then
             cboPageNo.Text = cboPageNo.List(cboPageNo.ListIndex + 1)
        End If
    End If
End Sub



Private Sub setPageButtons()
    If currTotalPages = 1 Then
        cmdPrevPage.Enabled = False
        cmdNextPage.Enabled = False
    Else
        If Val(cboPageNo.Text) = 1 Then
             cmdPrevPage.Enabled = False
             cmdNextPage.Enabled = True
        ElseIf Val(cboPageNo.Text) = currTotalPages Then
             cmdPrevPage.Enabled = True
             cmdNextPage.Enabled = False
        Else
             cmdPrevPage.Enabled = True
             cmdNextPage.Enabled = True
        End If
    End If
End Sub



Private Sub HScroll1_Change()
   Select Case mSizeNo
      Case 1
          Pic1.Left = -HScroll1.Value
      Case 2
          Pic2.Left = -HScroll1.Value
      Case 3
          Pic3.Left = -HScroll1.Value
      Case 4
          Pic4.Left = -HScroll1.Value
      Case 5
          Pic5.Left = -HScroll1.Value
   End Select
End Sub



Private Sub VScroll1_Change()
   Select Case mSizeNo
      Case 1
          Pic1.Top = -VScroll1.Value
      Case 2
          Pic2.Top = -VScroll1.Value
      Case 3
          Pic3.Top = -VScroll1.Value
      Case 4
          Pic4.Top = -VScroll1.Value
      Case 5
          Pic5.Top = -VScroll1.Value
   End Select
End Sub



Private Sub HScroll1_Scroll()
   HScroll1_Change
End Sub



Private Sub VScroll1_Scroll()
   VScroll1_Change
End Sub



Private Sub ChangePreview()
   Pic1.Visible = False
   Pic2.Visible = False
   Pic3.Visible = False
   Pic4.Visible = False
   Pic5.Visible = False
   Select Case mSizeNo
      Case 1
          Pic1.Visible = True
          Pic1.Left = 0
          setScrollMax Pic1
      Case 2
          Pic2.Visible = True
          Pic2.Left = (PicZ.Width - Pic2.Width) / 2
          setScrollMax Pic2
      Case 3
          Pic3.Visible = True
          Pic3.Left = (PicZ.Width - Pic3.Width) / 2
          setScrollMax Pic3
      Case 4
          Pic4.Visible = True
          Pic4.Left = (PicZ.Width - Pic4.Width) / 2
          setScrollMax Pic4
      Case 5
          Pic5.Visible = True
          Pic5.Left = (PicZ.Width - Pic5.Width) / 2
          setScrollMax Pic5
   End Select
End Sub



' With our chosen Style, Combo does not honor "Change", use "Click" instead
Private Sub cboScale_Click()
    cmdZoomIn.Enabled = True
    cmdZoomOut.Enabled = True
    mSizeNo = cboScale.ListIndex + 1
    If mSizeNo = 1 Then
         cmdZoomIn.Enabled = False
    ElseIf mSizeNo = 5 Then
         cmdZoomOut.Enabled = False
    End If
    ChangePreview
End Sub



Private Sub cmdPrint_Click()
     gprint = True
     Unload Me
End Sub



Private Sub cmdZoomin_click()
     If mSizeNo = 1 Then
          Exit Sub
     End If
     mSizeNo = mSizeNo - 1
     cboScale.Text = cboScale.List(mSizeNo - 1)
     cmdZoomIn.Enabled = True
     cmdZoomOut.Enabled = True
     If mSizeNo = 1 Then
          cmdZoomIn.Enabled = False
     ElseIf mSizeNo = 5 Then
          cmdZoomOut.Enabled = False
     End If
End Sub



Private Sub cmdzoomout_click()
     If mSizeNo = 5 Then
         Exit Sub
     End If
     mSizeNo = mSizeNo + 1
     cboScale.Text = cboScale.List(mSizeNo - 1)
     cmdZoomIn.Enabled = True
     cmdZoomOut.Enabled = True
     If mSizeNo = 1 Then
          cmdZoomIn.Enabled = False
     ElseIf mSizeNo = 5 Then
          cmdZoomOut.Enabled = False
     End If
End Sub



Private Function MakeSizes(ByVal mofSize As Integer) As Boolean
    Dim SrcX As Long, SrcY As Long
    Dim DestX As Long, DestY As Long
    Dim SrcWidth As Long, SrcHeight As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim SrcHDC As Long, DestHDC As Long
    Dim mresult
      
    SrcX = 0: SrcY = 0: DestX = 0: DestY = 0
    SrcWidth = PicX.ScaleWidth
    SrcHeight = PicX.ScaleHeight
    SrcHDC = PicX.hdc
    Select Case mofSize
       Case 1
          DestWidth = Pic1.ScaleWidth
          DestHeight = Pic1.ScaleHeight
          DestHDC = Pic1.hdc
          
      Case 2
          DestWidth = Pic2.ScaleWidth
          DestHeight = Pic2.ScaleHeight
          DestHDC = Pic2.hdc
       
      Case 3
          DestWidth = Pic3.ScaleWidth
          DestHeight = Pic3.ScaleHeight
          DestHDC = Pic3.hdc
          
      Case 4
          DestWidth = Pic4.ScaleWidth
          DestHeight = Pic4.ScaleHeight
          DestHDC = Pic4.hdc
      Case 5
          DestWidth = Pic5.ScaleWidth
          DestHeight = Pic5.ScaleHeight
          DestHDC = Pic5.hdc
    End Select

    mresult = StretchBlt(DestHDC, DestX, DestY, DestWidth, DestHeight, SrcHDC, _
      SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy)

    If mresult = 0 Then
        MsgBox "Error occurred in sizing images. Cannot continue"
        MakeSizes = False
    Else
        MakeSizes = True
    End If
End Function



Private Sub cmdClose_Click()
    Unload Me
End Sub



' To display the same as it would print on the selected printer
Function DocWYSIWYG(RTB As Control) As Long
     Dim LeftOffSet As Long
     Dim LeftMargin As Long, RightMargin As Long
     Dim linewidth As Long
     Dim PrinterhDC As Long
     Dim r As Long
     Printer.ScaleMode = vbTwips
     
               ' Get the offset to the printable area on the page in twips
          LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
             112), vbPixels, vbTwips)
             
     LeftMargin = gLeft - LeftOffSet 'Margin * 1440
     RightMargin = (Printer.Width - gRight) - LeftOffSet 'Margin * 1440
     linewidth = RightMargin - LeftMargin
     DocWYSIWYG = linewidth
  
End Function




Sub PrintPreviewPage(inControl As Control, InPage As Integer)
    Dim PageCtn
    Dim LeftOffSet As Long
    Dim TopOffSet As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    
    inControl.Picture = LoadPicture()
    
      ' Get the offsett to the printable area on the page in twips
      LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         112), vbPixels, vbTwips)
      TopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         113), vbPixels, vbTwips)
         
      ' Calculate the Left, Top, Right, and Bottom margins
      LeftMargin = gLeft - LeftOffSet
      TopMargin = gTop - TopOffSet
      RightMargin = (Printer.Width - gRight) - LeftOffSet
      BottomMargin = (Printer.Height - gBottom) - TopOffSet

      ' Set printable area rect. Note in frmPrintPreview, scaleModes are all in pixels,
      ' have to compute the twips equivalent
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = Printer.ScaleWidth
    rectPage.Bottom = Printer.ScaleHeight
   ' rectPage.Right = inControl.Width * Screen.TwipsPerPixelX
   ' rectPage.Bottom = inControl.Height * Screen.TwipsPerPixelY
 
      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = LeftMargin 'gLeft - 0 'Margin * 1440
    rectDrawTo.Top = TopMargin   ''Margin * 1440
    rectDrawTo.Right = RightMargin   'inControl.Width * Screen.TwipsPerPixelX - (gRight - 360) 'Margin * 1440
    rectDrawTo.Bottom = BottomMargin   'inControl.Height * Screen.TwipsPerPixelY - (gBottom - 0) 'Margin * 1440
 
    mFormatRange.hdc = inControl.hdc           ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = inControl.hdc     ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo       ' Area on page to draw to
    mFormatRange.rectPage = rectPage           ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0      ' Start of text
    mFormatRange.mCharRange.lastChar = -1      ' End of the text
    
    TextLength = Len(frmMainText.Text1.Text)


    PageCtn = 1
    Do
        newStartPos = SendMessage(frmMainText.Text1.hwnd, EM_FORMATRANGE, True, mFormatRange)
        If newStartPos >= TextLength Then
            Exit Do
        End If
        If PageCtn = InPage Then
            Exit Do
        End If
        inControl.Picture = LoadPicture()
        mFormatRange.mCharRange.firstChar = newStartPos    ' Starting position for next page
        mFormatRange.hdc = inControl.hdc
        mFormatRange.hdcTarget = inControl.hdc
        PageCtn = PageCtn + 1
        DoEvents
    Loop
    dumpaway = SendMessage(inControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub



' Test how many pages are there in total
Function PageCtnProc(inControl As Control) As Integer
    Dim mPageCtn As Integer
    Dim LeftOffSet As Long
    Dim TopOffSet As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
      ' Get the offsett to the printable area on the page in twips
      LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         112), vbPixels, vbTwips)
      TopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         113), vbPixels, vbTwips)
         
      ' Calculate the Left, Top, Right, and Bottom margins
      LeftMargin = gLeft - LeftOffSet
      TopMargin = gTop - TopOffSet
      RightMargin = (Printer.Width - gRight) - LeftOffSet
      BottomMargin = (Printer.Height - gBottom) - TopOffSet

      ' Set printable area rect. Note in frmPrintPreview, scaleModes are all in pixels,
      ' have to compute the twips equivalent
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = Printer.ScaleWidth
    rectPage.Bottom = Printer.ScaleHeight
    
      ' Set printable area rect.
    'rectPage.Left = 0
    'rectPage.Top = 0
    'rectPage.Right = inControl.Width * Screen.TwipsPerPixelX
    'rectPage.Bottom = inControl.Height * Screen.TwipsPerPixelY
    
      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = LeftMargin  'gLeft - 0 'Margin * 1440
    rectDrawTo.Top = TopMargin   ''Margin * 1440
    rectDrawTo.Right = RightMargin   'inControl.Width * Screen.TwipsPerPixelX - (gRight - 360) 'Margin * 1440
    rectDrawTo.Bottom = BottomMargin
 
      ' Set rect in which to print (relative to printable area)
   ' rectDrawTo.Left = gLeft - 0 'Margin * 1440
   ' rectDrawTo.Top = gTop - 360 'Margin * 1440
   ' rectDrawTo.Right = inControl.Width * Screen.TwipsPerPixelX _
   '      - (gRight - 360) 'Margin * 1440
   ' rectDrawTo.Bottom = inControl.Height * Screen.TwipsPerPixelY _
   '      - (gBottom - 0) 'Margin * 1440
 
      ' Set up the print instructions
    mFormatRange.hdc = inControl.hdc           ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = inControl.hdc     ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo       ' Area on page to draw to
    mFormatRange.rectPage = rectPage           ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0      ' Start of text
    mFormatRange.mCharRange.lastChar = -1      ' End of the text

    TextLength = Len(frmMainText.Text1.Text)

    mPageCtn = 1
    Do
          ' Print the page by sending EM_FORMATRANGE message
        newStartPos = SendMessage(frmMainText.Text1.hwnd, EM_FORMATRANGE, True, mFormatRange)
        If newStartPos >= TextLength Then
            Exit Do
        End If
        mFormatRange.mCharRange.firstChar = newStartPos       ' Starting position for next page
        mFormatRange.hdc = inControl.hdc
        mFormatRange.hdcTarget = inControl.hdc
        
        mPageCtn = mPageCtn + 1
        DoEvents
    Loop
    
    inControl.Picture = LoadPicture()
    dumpaway = SendMessage(inControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    PageCtnProc = mPageCtn
End Function



Sub DocPrintProc()

'This routine is not used in this version.  Instead we
'print with the main forms print routine which is based on
'Microsoft Knowledgebase WYSIWYG article

    On Error Resume Next
    DoEvents
    frmPrintPreview.picP.Picture = LoadPicture()
    
    Dim gcdg As Object
    Dim mFromPage As Integer, mToPage As Integer, mpage As Integer
    Dim mSelective As Boolean
    Dim mHwnd As Long
    Dim LeftOffSet As Long
    Dim TopOffSet As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    
    Set gcdg = frmMainText.cmDlg
    gcdg.DialogTitle = "Print"
    gcdg.CancelError = True

    gcdg.flags = 0
    gcdg.flags = cdlPDReturnDC
      ' Allow user select page range
    gcdg.Min = 1                           ' Must set to work around MS bug
    gcdg.Max = currTotalPages
    gcdg.FromPage = 1                      ' To force open up page selection
    gcdg.ToPage = currTotalPages
    
    If Len(frmMainText.Text1.SelText) > 0 Then
         gcdg.flags = gcdg.flags + cdlPDSelection + cdlPDPageNums
    Else
         gcdg.flags = gcdg.flags + cdlPDNoSelection + cdlPDPageNums
    End If

    gcdg.ShowPrinter
    
    If Err = MSComDlg.cdlCancel Then
         Exit Sub
    End If
    
    mSelective = False
    If (gcdg.flags And cdlPDSelection) <> 0 Then
         If Len(frmMainText.Text1.SelText) = 0 Then
              MsgBox "No selected text yet"
              Exit Sub
         End If
         mSelective = True
    End If
   
    If frmMainText.WindowState <> 1 Then
    Else
        MsgBox "Cannot proceed with minimized screen"
        Exit Sub
    End If
    
    'If MsgBox("Proceed to print", vbYesNo + vbQuestion) = vbNo Then
    '    Exit Sub
    'End If
    
    DocWYSIWYG frmMainText.Text1
    'frmMainText.Move 0, 0
    
    mHwnd = frmMainText.Text1.hwnd            ' Assume first

    If (gcdg.flags And cdlPDPageNums) <> 0 Then
        mFromPage = gcdg.FromPage
        mToPage = gcdg.ToPage
    Else
        mFromPage = 1
        mToPage = currTotalPages
          ' If print selection only, transcribe selected contents to textHidden for print
         '-----------------------------------
          ' Note that the SelPrint sends only the selected text to the target device, hence
          ' the following would otherwise be OK if not because (i) it does not let us have
          ' a control over the print margins and (ii) it does not allow us to include
          ' pictures.
          'If (gcdg.Flags And cdlPDSelection) <> 0 Then
          '    Printer.Print ""
          '    frmMainText.Text1.SelPrint gcdg.hdc
          '    Exit Sub
          'End If
         '-----------------------------------
        If mSelective Then
            frmMainText.txtInsert.Text = ""
            DocWYSIWYG frmMainText.txtInsert
               ' We could directly transcribe Seltext into textHidden, but then it covers
               ' text only - it would not cover a picture as well.  Therefore, we have to
               ' do the following workaround.
            frmMainText.picHidden.Picture = LoadPicture()
               ' Due to implications of having a picture in our printout, each time we
               ' allow one page only.
            frmMainText.picHidden.Width = frmMainText.Text1.Width
            frmMainText.picHidden.Height = frmMainText.Text1.Height
            frmMainText.Text1.SelPrint frmMainText.picHidden.hdc
            frmMainText.picHidden.Picture = frmMainText.picHidden.Image
            
            Clipboard.Clear
            Clipboard.SetData frmMainText.picHidden.Picture
            SendMessage frmMainText.txtInsert.hwnd, WM_PASTE, 0, 0
            
            mHwnd = frmMainText.txtInsert.hwnd        ' We change earlier value
        End If
    End If
    
    Printer.Print ""
    Printer.ScaleMode = vbTwips
    
      ' Get the offsett to the printable area on the page in twips
      LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         112), vbPixels, vbTwips)
      TopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         113), vbPixels, vbTwips)

      ' Calculate the Left, Top, Right, and Bottom margins
      LeftMargin = gLeft - LeftOffSet
      TopMargin = gTop - TopOffSet
      RightMargin = (Printer.Width - gRight) - LeftOffSet
      BottomMargin = (Printer.Height - gBottom) - TopOffSet
    
      ' Set printable rect area
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = Printer.ScaleWidth
    rectPage.Bottom = Printer.ScaleHeight

      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = LeftMargin  'gLeft - 0 'Margin * 1440
    rectDrawTo.Top = TopMargin   'gTop - 360 'Margin * 1440
    rectDrawTo.Right = RightMargin   'Printer.ScaleWidth - gRight - 360 'Margin * 1440
    rectDrawTo.Bottom = BottomMargin   'Printer.ScaleHeight - gBottom - 0 'Margin * 1440

     ' Dump earlier pages if any to PicP before reaching first wanted page
    mFormatRange.hdc = frmPrintPreview.picP.hdc
    mFormatRange.hdcTarget = frmPrintPreview.picP.hdc
    
    newStartPos = 0                                   ' Next char to start
    mFormatRange.rectRegion = rectDrawTo              ' Area on page to draw to
    mFormatRange.rectPage = rectPage                  ' Entire size of page
    mFormatRange.mCharRange.firstChar = newStartPos   ' Start of text
    mFormatRange.mCharRange.lastChar = -1             ' End of the text

    If Not mSelective Then
        TextLength = Len(frmMainText.Text1.Text)
    Else
        TextLength = Len(frmMainText.txtInsert.Text)
    End If
    
      ' Dumping if any
    mpage = 1
    Do
        If mpage = mFromPage Then
            Exit Do
        End If
        
        ' Don't clear picture box here, unless you want to print from first page always.
        
          ' Print the page by sending EM_FORMATRANGE message
        newStartPos = SendMessage(mHwnd, EM_FORMATRANGE, True, mFormatRange)
        If newStartPos >= TextLength Then
            Exit Do
        End If
        mFormatRange.mCharRange.firstChar = newStartPos    ' Starting position for next page
        mFormatRange.hdc = frmPrintPreview.picP.hdc
        mFormatRange.hdcTarget = frmPrintPreview.picP.hdc
        mpage = mpage + 1
        DoEvents
    Loop

       ' Must cleanse memory here before print, otherwise font will not be right
    dumpaway = SendMessage(mHwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    
    If newStartPos >= TextLength Then
        Exit Sub
    End If
    
       ' Have to reinitialize printer here
    Printer.Print ""
    Printer.ScaleMode = vbTwips
    
       ' Actual print to printer, starting from the user-selected Page No.
    mFormatRange.hdc = Printer.hdc
    mFormatRange.hdcTarget = Printer.hdc
    
      ' Update char range
    mFormatRange.mCharRange.firstChar = newStartPos
    
    Do
        newStartPos = SendMessage(mHwnd, EM_FORMATRANGE, True, mFormatRange)
        If newStartPos >= TextLength Then
            Exit Do
        End If
        If mpage >= mToPage Then
            Exit Do
        End If
        mFormatRange.mCharRange.firstChar = newStartPos
        Printer.NewPage                  ' Move on to next page
        Printer.Print ""                 ' Re-initialize hDC
        mFormatRange.hdc = Printer.hdc
        mFormatRange.hdcTarget = Printer.hdc
        mpage = mpage + 1
        DoEvents
    Loop
      ' Commit the print job
    Printer.EndDoc
      ' Free up memory
    dumpaway = SendMessage(mHwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    frmMainText.txtInsert.Text = ""
    frmMainText.picHidden.Picture = LoadPicture()
End Sub




Private Sub setScrollMax(inPic As PictureBox)
    HScroll1.Max = inPic.ScaleWidth - PicZ.ScaleWidth
    VScroll1.Max = inPic.ScaleHeight - PicZ.ScaleHeight
    If HScroll1.Max <= 0 Then
         HScroll1.Max = 0
    Else
         If PicZ.ScaleWidth / HScroll1.Max < 1 Then
              HScroll1.SmallChange = 1
         Else
              HScroll1.SmallChange = PicZ.ScaleWidth / HScroll1.Max
         End If
         HScroll1.LargeChange = HScroll1.SmallChange
         If HScroll1.Max >= 40 Then
              HScroll1.LargeChange = HScroll1.Max / 20
              If HScroll1.LargeChange < HScroll1.SmallChange Then
                   HScroll1.LargeChange = HScroll1.SmallChange
              End If
         End If
    End If
    If VScroll1.Max <= 0 Then
         VScroll1.Max = 0
    Else
         If PicZ.ScaleHeight / VScroll1.Max < 1 Then
              VScroll1.SmallChange = 1
         Else
              VScroll1.SmallChange = PicZ.ScaleHeight / VScroll1.Max
         End If
         VScroll1.LargeChange = VScroll1.SmallChange
         If VScroll1.Max >= 40 Then
              VScroll1.LargeChange = VScroll1.Max / 20
              If VScroll1.LargeChange < VScroll1.SmallChange Then
                   VScroll1.LargeChange = VScroll1.SmallChange
              End If
         End If
    End If
End Sub

