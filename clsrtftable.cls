VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRTFtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCell"
Attribute VB_Ext_KEY = "Member1" ,"clsColumn"
Attribute VB_Ext_KEY = "Member2" ,"clsRow"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'   clsTable - used to insert a table into a RichTextBox

'   Copyright (C) 2005  William Moeur
'   This program is free software; you can redistribute it and/or modify it

'Properties - all sizes are in twips
'   xLeft - Position of the left edge of the table
'   isCentered - Set to True to center the table
'   Rows - Sets or returns the number of rows in the table
'   Columns - Sets or returns the number of columns in the table
'   Row - An Array of Rows (1 to Rows)
'   Column - An Array of columns (1 to Columns)
'       Column(i).xWidth - Width of the ith column
'   Cell - A 2-d Array of Cells (1 to Rows, 1 to Columns)
'       Cell(r, c).Contents - Sets or returns the contents of the cell

'Methods
'   InsertTable(RTB As RichTextBox) - Inserts the table into the RichTextBox
'                                       at the currrent cursor position.


'local variable(s) to hold property value(s)
Private mvarclsColumn() As clsColumn
Private mvarclsRow() As clsRow
Private mvarclsCell() As clsCell
Private mvarisCentered As Boolean
Private mvarxLeft As Long
Private mvarColumns As Integer
Private mvarRows As Integer

Public Property Let xLeft(ByVal vData As Long)
    mvarxLeft = vData
End Property

Public Property Get xLeft() As Long
    xLeft = mvarxLeft
End Property

Public Property Let isCentered(ByVal vData As Boolean)
    mvarisCentered = vData
End Property

Public Property Get isCentered() As Boolean
    isCentered = mvarisCentered
End Property

Public Property Get Cell(RowIndex As Integer, ColIndex As Integer) As clsCell
    Set Cell = mvarclsCell(RowIndex, ColIndex)
End Property

Public Property Set Cell(RowIndex As Integer, ColIndex As Integer, vData As clsCell)
    Set mvarclsCell(RowIndex, ColIndex) = vData
End Property

Public Property Get Row(Index As Integer) As clsRow
    Set Row = mvarclsRow(Index)
End Property

Public Property Set Row(Index As Integer, vData As clsRow)
    Set mvarclsRow(Index) = vData
End Property

Public Property Get Column(Index As Integer) As clsColumn
    Set Column = mvarclsColumn(Index)
End Property

Public Property Set Column(Index As Integer, vData As clsColumn)
    Set mvarclsColumn(Index) = vData
End Property

Private Sub Class_Initialize()
    mvarxLeft = 36
    
End Sub

Private Sub Class_Terminate()

    RemoveCells
    RemoveColumns
    RemoveRows

End Sub

Public Property Let Rows(ByVal vData As Integer)
'sets number of rows and creates row array
Dim i As Integer
    RemoveRows
    mvarRows = vData
    ReDim mvarclsRow(1 To vData)
    For i = 1 To vData
        Set mvarclsRow(i) = New clsRow
    Next i
    AddCells
End Property

Public Property Get Rows() As Integer
    Rows = mvarRows
End Property

Public Property Let Columns(ByVal vData As Integer)
'sets number of columns
Dim i As Integer
    RemoveColumns
    mvarColumns = vData
    ReDim mvarclsColumn(1 To vData)
    For i = 1 To vData
        Set mvarclsColumn(i) = New clsColumn
    Next i
    AddCells
End Property

Public Property Get Columns() As Integer
    Columns = mvarColumns
End Property

Public Sub InsertTable(RTB As RichTextBox)
Dim strRTF As String
Dim strInsert As String
Dim r As Integer, c As Integer, i As Integer, w As Long

'-- table specs
'\trowd - table row defaults
'\trgaph108 - space between cells
'\trleft36 - left side of table
'\trqc - center table
'\cellx1636 - column 1 width
'\cellx3236 - column 2 width
'\cellx4836 - column 3 width

'-- Begin Table
'\pard

'-- Row 1
'\intbl <text>\cell <text>\cell <text>\cell\row

'-- Row 2
'\intbl <text>\cell <text>\cell <text>\cell\row

'-- end of table
'\pard
    
    If mvarRows = 0 Or mvarColumns = 0 Then Exit Sub

    strInsert = "\par \trowd"
    If mvarisCentered Then strInsert = strInsert & "\trqc"
    strInsert = strInsert & "\trgaph108\trleft" & CStr(mvarxLeft)
    'set column widths
    For c = 1 To mvarColumns
        strInsert = strInsert & "\cellx"
        w = mvarxLeft
        For i = 1 To c
          w = w + mvarclsColumn(i).xWidth
        Next i
        strInsert = strInsert & CStr(w)
    Next c
    'begin table
    strInsert = strInsert & "\pard"
    'do each row
    Dim strRow As String
    For r = 1 To mvarRows
        strRow = "\intbl"
        For c = 1 To mvarColumns
            strRow = strRow & " " & CStr(Cell(r, c).Contents) & "\cell"
        Next c
        strInsert = strInsert & strRow & "\row" & vbCrLf
    Next r
    strInsert = strInsert & "\pard" & vbCrLf & "\par" & vbCrLf
    'Debug.Print strInsert

    With RTB
        'mark current insertion point
        .SelText = Chr(&H80)
        'get current rtf text
        strRTF = .TextRTF
        'insert out text
        strRTF = Replace(.TextRTF, "\'80", strInsert)
        .TextRTF = strRTF
    End With

End Sub

Private Sub RemoveColumns()
Dim i As Integer
    For i = 1 To mvarColumns
        Set mvarclsColumn(i) = Nothing
    Next i
    mvarColumns = 0
End Sub

Private Sub RemoveRows()
Dim i As Integer
    For i = 1 To mvarRows
        Set mvarclsRow(i) = Nothing
    Next i
    mvarRows = 0
End Sub

Private Sub RemoveCells()
Dim r As Integer, c As Integer
If mvarRows = 0 Or mvarColumns = 0 Then Exit Sub
On Error Resume Next
    For r = 1 To mvarRows
        For c = 1 To mvarColumns
            Set mvarclsCell(r, c) = Nothing
        Next c
    Next r
End Sub

Private Sub AddCells()
Dim r As Integer, c As Integer
If mvarRows = 0 Or mvarColumns = 0 Then Exit Sub
    RemoveCells
    ReDim mvarclsCell(1 To mvarRows, 1 To mvarColumns) As clsCell

    For r = 1 To mvarRows
        For c = 1 To mvarColumns
            Set mvarclsCell(r, c) = New clsCell
        Next c
    Next r
End Sub
