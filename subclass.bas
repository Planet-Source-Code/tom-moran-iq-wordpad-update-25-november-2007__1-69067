Attribute VB_Name = "modSubClass"
'******************************************************************************************
'            Module for setting the position, width, and height of a ComboList
'******************************************************************************************
Option Explicit
'******************************************************************************************
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'---------------------------------------------------------
Private Const GWL_WNDPROC& = (-4)
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_ERASEBKGND& = &H14
Private Const WM_PAINT& = &HF
Private Const CB_GETCOUNT& = &H146
Private Const CB_GETITEMHEIGHT& = &H154
'************************************************************************
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd&, lpRect As Rect)
'---------------------------------------------------------
Public Const CB_SHOWDROPDOWN& = &H14F
Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" _
               (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
'************************************************************************
Private OldProc&              '# Address of the original window procedure
'---------------------------------------------------------
Private l_width&, l_height&   '# Width and height of the combolist in Pixel
Private cl_left&, cl_top&     '# Left- and Top-Position of the list in Pixel
Private SH&                   '# Screen Height in Pixel
Private brect As Rect         '# current rectangle of the combobox
Private bDiff&                '# difference in width between box and list
Private C_hwnd&               '# Handle of the combobox
Private cl_hwnd&              '# Handle of combolist
Private Alignment&            '# alignment of combolist
Private Adapt As Boolean      '# If True, adapt the list height to the entry number
'******************************************************************************************
'# Here the real work will be done ...
Public Function WindowProc&(ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Static cl_rect As Rect, cl_width&, cl_height&, c_count&, i_height&
'Debug.Print Hex(wMsg)

Select Case wMsg
'# Watch the message WM_CTLCOLORLISTBOX for getting the handle of the combolist.
Case WM_CTLCOLORLISTBOX
  If cl_hwnd = 0 Then cl_hwnd = lParam

'# Watch the message WM_ERASEBKGND for positioning the combolist.
Case WM_ERASEBKGND, WM_PAINT
  Call GetWindowRect(cl_hwnd, cl_rect)  '# get the current position of the combolist
  With cl_rect
    cl_width = .Right - .Left
    cl_height = .Bottom - .Top
    If l_width <> 0 Then cl_width = l_width
    If l_height <> 0 Then
      If SH = 0 Then                    '# get the screen height, if not yet done
        With Screen
          SH = .Height \ .TwipsPerPixelY
        End With
      End If
      If Adapt Then                     '# adapt the list height to the entry number
        c_count = SendMessage(C_hwnd, CB_GETCOUNT, 0, ByVal 0&)
        i_height = SendMessage(C_hwnd, CB_GETITEMHEIGHT, 0, ByVal 0&)
        cl_height = c_count * i_height + 2 '# We must add "2" for the bounding rectangle
      Else
        cl_height = l_height
      End If
      '# Check whether the lower edge of the list would
      '# exceed the screen; if so, delimit the list height
      With cl_rect
        If .Top + cl_height > SH Then
          cl_height = SH - .Top
        End If
      End With
    End If
  End With
  
  Call GetWindowRect(C_hwnd, brect)     '# get the current position of the combo
  With brect
    cl_top = .Bottom                    '# align the top of the list to the bottom of the combobox
    Select Case Alignment
    Case 0: cl_left = .Left             '# left aligned
    Case 1: cl_left = .Left - bDiff     '# right aligned
    Case 2: cl_left = .Left - bDiff \ 2 '# centered
    End Select
  End With
  '# position the list
  Call SetWindowPos(cl_hwnd, 0, cl_left, cl_top, cl_width, cl_height, 0)
End Select

WindowProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
End Function

'******************************************************************************************
Sub SubClass(ByVal Chwnd&, ByVal YesNo As Boolean, Optional LW& = 0, Optional LH& = 0, Optional Align% = 0)
Dim rct As Rect

Select Case YesNo
Case True
  '# Check whether we should adapt the list height to the number of entries
  Adapt = IIf(LH = -1, True, False)
  '# Subclassing of the ComboBox
  OldProc = SetWindowLong(Chwnd, GWL_WNDPROC, AddressOf WindowProc)
  '# get the parameter
  l_width = LW: l_height = LH
  Alignment = Align
  Call GetWindowRect(Chwnd, rct)        '# get the rectangle of the ComboCBox
  With rct
    bDiff = l_width - (.Right - .Left)  '# store the difference of width for alignment
  End With
  C_hwnd = Chwnd

Case False
  '# restore the original window proc
  Call SetWindowLong(Chwnd, GWL_WNDPROC, OldProc)
End Select
End Sub

'******************************************************************************************
Sub SetAlignment(ByVal ALM&)
  Alignment = ALM
End Sub

'******************************************************************************************

