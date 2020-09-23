Attribute VB_Name = "modPrinter"

Option Explicit

   Private Type Rect
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
   End Type

   Private Type CharRange
     cpMin As Long     ' First character of range (0 for start of doc)
     cpMax As Long     ' Last character of range (-1 for end of doc)
   End Type

   Private Type FormatRange
     hdc As Long       ' Actual DC to draw on
     hdcTarget As Long ' Target DC for determining text formatting
     rc As Rect        ' Region of the DC to draw to (in twips)
     rcPage As Rect    ' Region of the entire DC (page size) (in twips)
     chrg As CharRange ' Range of text to draw (see above declaration)
   End Type

   Public Const WM_USER As Long = &H400
   Private Const EM_FORMATRANGE As Long = WM_USER + 57
   Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
   Private Const PHYSICALOFFSETX As Long = 112
   Private Const PHYSICALOFFSETY As Long = 113
   
   Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hdc As Long, ByVal nIndex As Long) As Long
   Private Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" _
      (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, _
      lp As Any) As Long
   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
      (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
      ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

'Some public variables for printer/screen formatting

   Public gLeft As Long
   Public gRight As Long
   Public gTop As Long
   Public gBottom As Long
   Public gHeader As String
   Public gFooter As String
   Public gAlign As Integer '0=left 1=center 2=right
   Public gHeadAlign As Integer '0=left 1=center 2=right - Alignment for header
   Public gFootAlign As Integer '0=left 1=center 2=right - Alignment for footer
   Public gPageNumber As Long 'counts printed pages
   Public gOrientation As Integer
   Public gPaperSize As Variant
   Public gPaperWidth As Variant
   Public gPaperHeight As Variant
   Public gMargin As Long
  
   Private Prs As String 'The parsed header/footer string
   
   'for print preview
   'Public gLeftMargin As Integer
   'Public gRightMargin As Integer
   'Public gTopMargin As Integer
   'Public gBottomMargin As Integer
   Public gprint As Boolean
   

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '
   ' WYSIWYG_RTF - Sets an RTF control to display itself the same as it
   '               would print on the default printer
   '
   ' RTF - A RichTextBox control to set for WYSIWYG display.
   '
   ' LeftMarginWidth - Width of desired left margin in twips
   '
   ' RightMarginWidth - Width of desired right margin in twips
   '
   ' Returns - The length of a line on the printer in twips
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Public Function WYSIWYG_RTF(RTF As RichTextBox, _
       LeftMarginWidth As Long, _
      RightMarginWidth As Long, Enabled As Boolean) As Long
      Dim LeftOffset As Long, LeftMargin As Long, RightMargin As Long
      Dim linewidth As Long
      Dim PrinterhDC As Long
      Dim r As Long
      
      If Enabled = True Then
      
          ' Start a print job to initialize printer object
          Printer.Print Space(1)
          Printer.ScaleMode = vbTwips
    
          ' Get the offset to the printable area on the page in twips
          LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
             PHYSICALOFFSETX), vbPixels, vbTwips)
    
          ' Calculate the Left, and Right margins
          LeftMargin = LeftMarginWidth - LeftOffset
          RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    
          ' Calculate the line width
          linewidth = RightMargin - LeftMargin
    
          ' Create an hDC on the Printer pointed to by the Printer object
          ' This DC needs to remain for the RTF to keep up the WYSIWYG display
          PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)
    
          ' Tell the RTF to base it's display off of the printer
          '    at the desired line width
          r = SendMessage(RTF.hwnd, EM_SETTARGETDEVICE, PrinterhDC, _
             ByVal linewidth)
    
          ' Abort the temporary print job used to get printer info
          Printer.KillDoc

      WYSIWYG_RTF = linewidth
      
      Else
      
        'Turn OFF WYSWYG Display.
        SendMessageLong RTF.hwnd, EM_SETTARGETDEVICE, 0, 0
              
        WYSIWYG_RTF = RTF.Width
        
      End If
      
   End Function
Private Sub FooterPrint(LeftMargin As Long, BottomMargin As Long)
 Dim txtWidth As Long
 Dim MaxWidth As Long

 Prs = gFooter
 gAlign = 0
 MaxWidth = Printer.Width - LeftMargin * 2


'footer y location is constant
 Printer.CurrentY = BottomMargin + 360

 ParseHF Prs, gPageNumber

 txtWidth = Printer.TextWidth(Prs)

   Select Case gAlign
   
     Case 0 'left align
        Printer.CurrentX = LeftMargin

     Case 1 'center align
        Printer.CurrentX = (LeftMargin + (MaxWidth - txtWidth) / 2) - LeftMargin / 2
        
     Case 2 'right align
        Printer.CurrentX = (LeftMargin + (MaxWidth - txtWidth)) - LeftMargin / 2

        
  End Select
  
  Printer.Print Prs 'print the parsed footer
  
End Sub


Private Sub HeaderPrint(LeftMargin As Long, TopMargin As Long)
 Dim txtWidth As Long
 Dim MaxWidth As Long

 Prs = gHeader
 gAlign = 0
 MaxWidth = Printer.Width - LeftMargin * 2


'location is constant
 Printer.CurrentY = TopMargin - 360
 
 ParseHF Prs, gPageNumber
 
 txtWidth = Printer.TextWidth(Prs)
 

   Select Case gAlign
   
     Case 0 'left align
        Printer.CurrentX = LeftMargin

     Case 1 'center align
        Printer.CurrentX = (LeftMargin + (MaxWidth - txtWidth) / 2) - LeftMargin
        
     Case 2 'right align
        Printer.CurrentX = (LeftMargin + (MaxWidth - txtWidth)) - LeftMargin / 2

        
  End Select
  
  Printer.Print Prs 'print the parsed header


End Sub


Private Sub ParseHF(Prs$, gPageNumber)

'************************************************************
'This parses out requests for alignment of header/footer as
'well as request for page numbers, date, time filename
'************************************************************

Dim Pars As Integer

Pars = InStr(Prs$, "&p")
 If Pars Then
  Prs$ = Mid$(Prs$, 1, Pars - 1) & Str$(gPageNumber) & Mid$(Prs$, Pars + 2, Len(Prs$))
 End If
 
Pars = InStr(Prs$, "&f")
 If Pars Then
  Prs$ = Mid$(Prs$, 1, Pars - 1) & LastPart(IQFileName) & Mid$(Prs$, Pars + 2, Len(Prs$))
 End If
 
 Pars = InStr(Prs$, "&d")
  If Pars Then
   Prs$ = Mid$(Prs$, 1, Pars - 1) & Format(Now, "dddd, mmmm dd, yyyy") & Mid$(Prs$, Pars + 2, Len(Prs$))
  End If
  
 Pars = InStr(Prs$, "&t")
  If Pars Then
   Prs$ = Mid$(Prs$, 1, Pars - 1) & Format(Now, "h:mm AM/PM") & Mid$(Prs$, Pars + 2, Len(Prs$))
  End If

'Parse for alignment
 Pars = InStr(Prs$, "&r") 'right align
  If Pars Then
   gAlign = 2
   Prs$ = Mid$(Prs$, 1, Pars - 1) & Mid$(Prs$, Pars + 2, Len(Prs$))
  End If
  
 Pars = InStr(Prs$, "&c") 'center align
  If Pars Then
   gAlign = 1
   Prs$ = Mid$(Prs$, 1, Pars - 1) & Mid$(Prs$, Pars + 2, Len(Prs$))
  End If
  
 Pars = InStr(Prs$, "&l") 'center align
  If Pars Then
   gAlign = 0
   Prs$ = Mid$(Prs$, 1, Pars - 1) & Mid$(Prs$, Pars + 2, Len(Prs$))
  End If
 
End Sub

Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
      
'*****************************************************************************
'* This code largely from Microsoft KB - Modified by me for header/footers   *
'*****************************************************************************

On Error Resume Next

      Dim LeftOffset As Long, TopOffset As Long
      Dim LeftMargin As Long, TopMargin As Long
      Dim RightMargin As Long, BottomMargin As Long
      Dim fr As FormatRange
      Dim rcDrawTo As Rect
      Dim rcPage As Rect
      Dim TextLength As Long
      Dim NextCharPosition As Long
      Dim r As Long

      ' Start a print job to get a valid Printer.hDC
      
      Printer.Print " ";
      Printer.ScaleMode = vbTwips
      Printer.Font.Size = frmMainText.Font.Size
      

      DoEvents

      ' Get the offsett to the printable area on the page in twips
      LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETX), vbPixels, vbTwips)
      TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETY), vbPixels, vbTwips)

      ' Calculate the Left, Top, Right, and Bottom margins
      LeftMargin = LeftMarginWidth - LeftOffset
      TopMargin = TopMarginHeight - TopOffset
      RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
      BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

      ' Set printable area rect
      rcPage.Left = 0
      rcPage.Top = 0
      rcPage.Right = Printer.ScaleWidth
      rcPage.Bottom = Printer.ScaleHeight

      ' Set rect in which to print (relative to printable area)
      rcDrawTo.Left = LeftMargin
      rcDrawTo.Top = TopMargin
      rcDrawTo.Right = RightMargin
      rcDrawTo.Bottom = BottomMargin

      ' Set up the print instructions
      fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
      fr.hdcTarget = Printer.hdc  ' Point at printer hDC
      fr.rc = rcDrawTo            ' Indicate the area on page to draw to
      fr.rcPage = rcPage          ' Indicate entire size of page
      fr.chrg.cpMin = 0           ' Indicate start of text through
      fr.chrg.cpMax = -1          ' end of the text

      ' Get length of text in RTF
      TextLength = Len(RTF.Text)
      gPageNumber = 1

      ' Loop printing each page until done
      Do
        
        'Chk to see if we're printing header
        If Len(gHeader) Then
         Call HeaderPrint(LeftMargin, TopMargin)
        End If
           
         ' Print the page by sending EM_FORMATRANGE message
         NextCharPosition = SendMessage(RTF.hwnd, _
             EM_FORMATRANGE, True, fr)
             
        'Chk to see if we're printing footer
        If Len(gFooter) Then
         Call FooterPrint(LeftMargin, BottomMargin)
        End If
        
        'increment print page number
        gPageNumber = gPageNumber + 1
      
         If NextCharPosition >= TextLength Then Exit Do  'If done then exit
         fr.chrg.cpMin = NextCharPosition ' Starting position for next page
         Printer.NewPage                  ' Move on to next page
         Printer.Print Space(1) ' Re-initialize hDC
         fr.hdc = Printer.hdc
         fr.hdcTarget = Printer.hdc
      Loop

      ' Commit the print job
      Printer.EndDoc

      ' Allow the RTF to free up memory
      r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
      
   End Sub

