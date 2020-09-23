Attribute VB_Name = "modEnum"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module is just to close spell check anywhere if loaded
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MAX_PATH As Long = 260

Public Declare Function EnumWindows Lib "user32" _
   (ByVal lpEnumFunc As Long, _
    ByVal lParam As Long) As Long

Public Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" _
   (ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" _
    Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public ClosehWnd As Long
Public Const WM_CLOSE = &H10

Public Function EnumWindowProc(ByVal hwnd As Long, _
                               ByVal lParam As Long) As Long
   
  'working vars
   Dim sTitle As String
   
  'set up the strings to receive the class and
  'window text. You could use GetWindowTextLength,
  'but I'll cheat and use MAX_PATH instead.
   sTitle = Space$(MAX_PATH)
   Call GetWindowText(hwnd, sTitle, MAX_PATH)
   
   If LCase(Left$(sTitle, 20)) = "spell check anywhere" Then
    ClosehWnd = hwnd
    EnumWindowProc = 0
   End If
          
  'to continue enumeration, we must return True
  '(in C that's 1).  If we wanted to stop (perhaps
  'using if this as a specialized FindWindow method,
  'comparing a known class and title against the
  'returned values, and a match was found, we'd need
  'to return False (0) to stop enumeration. When 1 is
  'returned, enumeration continues until there are no
  'more windows left.
   EnumWindowProc = 1

End Function



