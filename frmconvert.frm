VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Word Document Conversion"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Converting Word Doc... Please wait"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   2475
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Sub DisableCloseWindowButton(frm As Form)
 Dim hSysMenu As Long
'Get the handle to this windows
'system menu
 hSysMenu = GetSystemMenu(frm.hwnd, 0)
'Remove the Close menu item
'This will also disable the close button
 RemoveMenu hSysMenu, 6, MF_BYPOSITION
'Lastly, we remove the seperator bar
 RemoveMenu hSysMenu, 5, MF_BYPOSITION
End Sub


Private Sub Form_Load()
 DisableCloseWindowButton Me
End Sub
