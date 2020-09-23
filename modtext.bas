Attribute VB_Name = "modText"
Option Explicit
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL As Long = &HB6

'**************************************
' For Fast Word Counting
'**************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

'Used to get Function results
Public There As Boolean
Public ret As Long


'main filename for program
Public IQFileName As String

'Globals for iQ WordPlus
Public PasteKey As Boolean
Public MyMsg As String
Public CurColumn As Long
Public ToolbarOn As Boolean
Public FormatbarOn As Boolean
Public StatusbarOn As Boolean
Public RulerOn As Boolean
Public WrapOn As Integer
Public i As Integer
Public fCancel As Boolean 'cancel flag
Public iQMeasurement As Integer
Public LineSpacing As Long

'for highlight feature
Public ColorHighLite As Long
Public HighColor As Boolean

'globals for tables
Public tRow As Integer
Public tCol As Integer
Public tWidth As Long
Public tColWidth(1 To 20)
Public tCenter As Boolean


Sub Main()
 On Error Resume Next
   ' Dim start As Double
    frmSplash.Show
    frmSplash.Refresh
    'start = Timer + 0.3
   ' While start > Timer
   ' Wend
    
    Load frmMainText
    Unload frmSplash
    frmMainText.Show
End Sub
Public Function GetColPos(tBox As Object) As Long
  GetColPos = tBox.SelStart - SendMessageByNum(tBox.hwnd, EM_LINEINDEX, -1&, 0&)
End Function
Public Function GetLineNum(tBox As Object) As Long
  GetLineNum = SendMessageByNum(tBox.hwnd, EM_LINEFROMCHAR, tBox.SelStart, 0&)
End Function
'not needed
'Public Function GetLineCount(tBox As RichTextBox) As Long
'  GetLineCount = SendMessageByNum(tBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
'End Function


Public Function LastPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  LastPart = Temp
 Else
  LastPart = Mid$(Temp, i + 1)
 End If
End Function
Public Function FirstPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  FirstPart = Temp
 Else
  FirstPart = Left$(Temp, i - 1)
 End If
End Function



Public Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function
Public Function WordCount(Text As String) As Long
    Dim dest() As Byte
    Dim i As Long


    If LenB(Text) Then
        ' Move the string's byte array into dest
        '     ()
        ReDim dest(LenB(Text))
        CopyMemory dest(0), ByVal StrPtr(Text), LenB(Text) - 1
        ' Now loop through the array and count t
        '     he words


        For i = 0 To UBound(dest) Step 2


            If dest(i) > 32 Then


                Do Until dest(i) < 33
                    i = i + 2
                Loop
                WordCount = WordCount + 1
            End If
        Next i
        Erase dest
    Else
        WordCount = 0
    End If
End Function
        

