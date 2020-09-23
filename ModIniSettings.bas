Attribute VB_Name = "ModIniSettings"
Option Explicit
' API functions used to read and write to INI.
' Used for handling the recent files list and options and print options.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Long

'Needed to keep full path of recent documents_
'which is different than parsed recent doc menu items
Public RecentDocs(0 To 8) As String
Public RecentFonts(0 To 8) As String

'Used to compact recent file paths-----------------
Private Const MAX_PATH As Long = 260

Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
   Alias "PathCompactPathExA" _
  (ByVal pszOut As String, _
   ByVal pszSrc As String, _
   ByVal cchMax As Long, _
   ByVal dwFlags As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long
'-----------------------------------------------------

'globals for ini options

Private TopForm As Long
Private LeftForm As Long
Private WidthForm As Long
Private HeightForm As Long
Private ToolB As String
Private StatusB As String
Private FormatB As String
Private WrapB As Integer
Private RulerB As String

Private retval As Long
Private key As String
Private l As Long
Private IniString As String

Public NameFont As String
Public SizeFont As Integer
Public BoldFont As Integer
Public ItalicFont As Integer
Public ColorFont As Long

Public Sub GetPrintOptions()

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)
  
  key = "PrintTop"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gTop = Val(Left$(IniString, l - 1))
    End If
  End If
  
    key = "PrintLeft"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gLeft = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "PrintRight"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gRight = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "PrintBottom"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gBottom = Val(Left$(IniString, l - 1))
    End If
  End If
  
  
  key = "PrintPaperSize"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gPaperSize = Val(Left$(IniString, l - 1))
    End If
  End If

  
  key = "PrintOrientation"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gOrientation = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "PaperWidth"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gPaperWidth = Val(Left$(IniString, l - 1))
    End If
  End If
    
 key = "PaperHeight"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gPaperHeight = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "PrintHeader"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gHeader = Left$(IniString, l - 1)
    End If
  End If
  
  key = "PrintFooter"
  retval = GetPrivateProfileString("Print Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      gFooter = Left$(IniString, l - 1)
    End If
  End If
  
  
 
End Sub


Public Sub GetRecentFonts()

'----------GetRecentFonts
  
  'Clear out the file array
  For i = 1 To 8
   RecentFonts(i) = ""
  Next
  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from iqWordPad.ini
  For i = 1 To 8
    key = "RecentFont" & i
    retval = GetPrivateProfileString("Recent Fonts", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
     
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the form's menu.
      RecentFonts(i) = TrimNull(IniString)
    End If
  Next i
  '===========================================================================
   'Now let's write just the font
    
   For i = 1 To 8
    If RecentFonts(i) = "" Then Exit For
    frmMainText.mnuRecentFonts(i).Caption = RecentFonts(i)
    frmMainText.mnuRecentFonts(i).Visible = True
   Next
    If RecentFonts(1) <> "" Then frmMainText.mnuRecentFonts(1).Enabled = True
End Sub

Public Function OnRecentFontsList(FontName As String) As Integer
  Dim j As Integer

  
  RecentFonts(0) = FontName
 
  For i = 1 To 8
    If RecentFonts(i) = FontName Then
     'move the recent font to the top
     For j = i To 1 Step -1
      RecentFonts(j) = RecentFonts(j - 1)
       ' Write changed file
      retval = WritePrivateProfileString("Recent Fonts", "RecentFont" & Trim((j)), RecentFonts(j), App.Path & "\iqWordPad.ini")
     Next j
     OnRecentFontsList = True
     Exit Function
    End If
  Next i
  
    OnRecentFontsList = False
End Function

Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function
Private Function MakeCompactedPathChrs(ByVal sPath As String, _
                                       ByVal cchMax As Long) As String

  'Truncates a path to a specified
  'number of characters by replacing
  'path components with ellipses.
   Dim ret As Long
   Dim buff As String
   
  'cchMax is the maximum number of characters
  'to be contained in the new string, **including
  'the terminating NULL character**. For example,
  'if cchMax = 8, the resulting string would contain
  'a maximum of 7 characters plus the termnating null.
  '
  'Because of this, we're add 1 to the value passed
  'as cchMax to ensure the resulting string is
  'the size requested.
   cchMax = cchMax + 1
   buff = Space$(MAX_PATH)
   ret = PathCompactPathEx(buff, sPath, cchMax, 0&)
   
   MakeCompactedPathChrs = TrimNull(buff)
   
End Function


Sub UpDateFileMenu(FileName As String)

        ' Check if OpenFileName is already on MRU list.
       
        retval = OnRecentFilesList(FileName)
        If retval = False Then
          ' Write OpenFileName to INI
          WriteRecentFiles (LCase(FileName))
          
        End If
        ' Update menus for most recent file list.
        GetRecentFiles
        
End Sub
Sub UpDateFontMenu(FontName As String)

        ' Check if OpenFileName is already on MRU list.
       
        retval = OnRecentFontsList(FontName)
        If retval = False Then
          ' Write OpenFileName to INI
          WriteRecentFonts FontName
          
        End If
        ' Update menus for most recent file list.
        GetRecentFonts
End Sub

Private Sub UpdateOptions()

If frmMainText.WindowState = 0 Then
 TopForm = frmMainText.Top
 LeftForm = frmMainText.Left
 WidthForm = frmMainText.Width
 HeightForm = frmMainText.Height
Else
 TopForm = (Screen.Height - 6700) / 2.4
 LeftForm = (Screen.Width - 10000) / 2
 WidthForm = 10000
 HeightForm = 0
End If

WrapB = WrapOn

If ToolbarOn = True Then
 ToolB = "On"
Else
 ToolB = "Off"
End If

If StatusbarOn = True Then
 StatusB = "On"
Else
 StatusB = "Off"
End If

If RulerOn = True Then
 RulerB = "On"
Else
 RulerB = "Off"
End If

If FormatbarOn = True Then
 FormatB = "On"
Else
 FormatB = "Off"
End If

End Sub

Public Sub WritePrintOptions()

  IniString = gTop
  retval = WritePrivateProfileString("Print Options", "PrintTop", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gLeft
  retval = WritePrivateProfileString("Print Options", "PrintLeft", IniString, App.Path & "\iqWordPad.ini")

  IniString = gRight
  retval = WritePrivateProfileString("Print Options", "PrintRight", IniString, App.Path & "\iqWordPad.ini")

  IniString = gBottom
  retval = WritePrivateProfileString("Print Options", "PrintBottom", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gPaperSize
  retval = WritePrivateProfileString("Print Options", "PrintPaperSize", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gPaperWidth
  retval = WritePrivateProfileString("Print Options", "PaperWidth", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gPaperHeight
  retval = WritePrivateProfileString("Print Options", "PaperHeight", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gOrientation
  retval = WritePrivateProfileString("Print Options", "PrintOrientation", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gHeader
  retval = WritePrivateProfileString("Print Options", "PrintHeader", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = gFooter
  retval = WritePrivateProfileString("Print Options", "PrintFooter", IniString, App.Path & "\iqWordPad.ini")
  

End Sub

Public Sub WriteRecentFiles(OpenFileName As String)
'=========== write recent files======================

  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 7 To 1 Step -1
    key = "RecentFile" & Trim(i)
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
    If retval And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFile" & Trim((i + 1))
      retval = WritePrivateProfileString("Recent Files", key, IniString, App.Path & "\iqWordPad.ini")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, App.Path & "\iqWordPad.ini")

End Sub
Public Function OnRecentFilesList(FileName As String) As Integer
  Dim j%
  
  RecentDocs(0) = FileName
 
  For i = 1 To 8
    'debug MsgBox LCase(RecentDocs(i)) & vbCrLf & LCase(Filename)
    If LCase(RecentDocs(i)) = LCase(FileName) Then
     'move the recent doc to the top
     For j = i To 1 Step -1
      RecentDocs(j) = RecentDocs(j - 1)
       ' Write changed file
      retval = WritePrivateProfileString("Recent Files", "RecentFile" & Trim((j)), RecentDocs(j), App.Path & "\iqWordPad.ini")
     Next j
     OnRecentFilesList = True
     Exit Function
    End If
  Next i
  
    OnRecentFilesList = False
End Function
Public Sub GetRecentFiles()

'----------GetRecentFiles
  
  'Clear out the file array
  For i = 1 To 8
   RecentDocs(i) = ""
  Next
  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from iqWordPad.ini
  For i = 1 To 8
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
     
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the form's menu.
      RecentDocs(i) = TrimNull(IniString)
    End If
  Next i
  '===========================================================================
   'Now let's write just the filename to the menu with shortened path
    
   For i = 1 To 8
    If RecentDocs(i) = "" Then Exit For
    frmMainText.mnuRecentFiles(i).Caption = MakeCompactedPathChrs(RecentDocs(i), 30)
    frmMainText.mnuRecentFiles(i).Visible = True
   Next
    If RecentDocs(1) <> "" Then frmMainText.mnuRecentFiles(1).Enabled = True
  '===============================================================================
End Sub

Public Function Exist(TF As String)
'Usage
 'XFileName = Command$
' There% = Exist(XFileName)
' If Not There% Then
'  MsgBox "Could not find " + XFileName
'  End
' Else
'  CommandView
' End If
Dim FileNum1 As Integer
Dim b As String
On Local Error GoTo Whoops
FileNum1 = FreeFile
 Open TF For Input As FileNum1
 Close FileNum1
 Exist = -1
 Exit Function
Whoops:
Exist = 0
b = TF + " file not found"
Exit Function

End Function
Sub GetOptions()

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)
  
  key = "TopForm"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      frmMainText.Top = Val(Left$(IniString, l - 1))
    End If
  End If
  
    key = "LeftForm"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      frmMainText.Left = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "WidthForm"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      frmMainText.Width = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "HeightForm"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
     If Val(Left$(IniString, l - 1)) > 2000 Then
      frmMainText.Height = Val(Left$(IniString, l - 1))
     Else
      frmMainText.Height = 6700
      frmMainText.WindowState = 2
     End If
    End If
  End If
  
  
  key = "ToolB"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ToolB = Left$(IniString, l - 1)
       If ToolB = "Off" Then
         frmMainText.xxxmnuTool.Checked = False
         frmMainText.picToolbar.Visible = False
         ToolbarOn = False
         DoEvents
       End If
    End If
  End If

  
  key = "WrapB"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      WrapOn = Val(Left$(IniString, l - 1))
    End If
  End If
  
  key = "StatusB"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      StatusB = Left$(IniString, l - 1)
       If StatusB = "Off" Then
         frmMainText.xxxmnuStatus.Checked = False
         frmMainText.StatusBar1.Visible = False
         StatusbarOn = False
       End If
    End If
  End If
  
  key = "RulerB"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      RulerB = Left$(IniString, l - 1)
       If RulerB = "Off" Then
         frmMainText.xxxmnuRuler.Checked = False
         frmMainText.RulerBar.Visible = False
         RulerOn = False
       End If
    End If
  End If
  
   key = "FormatB"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      FormatB = Left$(IniString, l - 1)
       If FormatB = "Off" Then
         frmMainText.xxxmnuFormatBar.Checked = False
         frmMainText.picFormatbar.Visible = False
         FormatbarOn = False
       End If
    End If
  End If
  
  key = "Measure"
  retval = GetPrivateProfileString("Options", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      iQMeasurement = Val(Left$(IniString, l - 1))
    End If
  End If
  
 End Sub

Public Sub WriteOptions()

  UpdateOptions
  
  IniString = TopForm
  retval = WritePrivateProfileString("Options", "TopForm", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = LeftForm
  retval = WritePrivateProfileString("Options", "LeftForm", IniString, App.Path & "\iqWordPad.ini")

  IniString = WidthForm
  retval = WritePrivateProfileString("Options", "WidthForm", IniString, App.Path & "\iqWordPad.ini")

  IniString = HeightForm
  retval = WritePrivateProfileString("Options", "HeightForm", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = WrapB
  retval = WritePrivateProfileString("Options", "WrapB", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = ToolB
  retval = WritePrivateProfileString("Options", "ToolB", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = StatusB
  retval = WritePrivateProfileString("Options", "StatusB", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = RulerB
  retval = WritePrivateProfileString("Options", "RulerB", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = FormatB
  retval = WritePrivateProfileString("Options", "FormatB", IniString, App.Path & "\iqWordPad.ini")
  
  IniString = iQMeasurement
  retval = WritePrivateProfileString("Options", "Measure", IniString, App.Path & "\iqWordPad.ini")
End Sub


Public Sub WriteRecentFonts(OpenFontName As String)

  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 7 To 1 Step -1
    key = "RecentFont" & Trim(i)
    retval = GetPrivateProfileString("Recent Fonts", key, "Not Used", IniString, Len(IniString), App.Path & "\iqWordPad.ini")
    If retval And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFont" & Trim((i + 1))
      retval = WritePrivateProfileString("Recent Fonts", key, IniString, App.Path & "\iqWordPad.ini")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Fonts", "RecentFont1", OpenFontName, App.Path & "\iqWordPad.ini")

End Sub


