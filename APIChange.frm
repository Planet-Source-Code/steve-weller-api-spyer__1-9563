VERSION 5.00
Begin VB.Form frmAPIChange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Window"
   ClientHeight    =   465
   ClientLeft      =   3030
   ClientTop       =   3330
   ClientWidth     =   2625
   Icon            =   "APIChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHwnd 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblHwndIs 
      Caption         =   "Window's &hWnd:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuChange 
      Caption         =   "&Change"
      Begin VB.Menu mnuChangeCaption 
         Caption         =   "&Caption"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuChangeVisible 
         Caption         =   "&Visible"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuChangeEnabled 
         Caption         =   "&Enabled"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuGet 
      Caption         =   "&Get"
      Begin VB.Menu mnuGethDC 
         Caption         =   "&hDC"
      End
      Begin VB.Menu mnuGetRect 
         Caption         =   "Window &Rect"
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "&Misc."
      Begin VB.Menu mnuMiscBringToTop 
         Caption         =   "&Bring To Top"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAPIChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' the window handle to use when
' changing window properties
Public WinHwnd As Long

' API declarations
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

' used for GetWindowRect
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2
Private Sub Form_Load()
txtHwnd = WinHwnd
End Sub





Private Sub Form_Unload(Cancel As Integer)
' reinitialize API Spyer form
frmAPISpy!tmrInfo.Enabled = True
SetWindowPos frmAPISpy.hwnd, IIf(frmAPISpy!chkOnTop.Value = vbChecked, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub mnuChangeCaption_Click()
' change the window's caption
Dim WndCaption$, R&, ChangeCaption$

' first get some info on the caption
WndCaption = Space$(256)
R = GetWindowText(WinHwnd, WndCaption, Len(WndCaption))

' ask user for new caption
ChangeCaption = InputBox$("Enter caption to change to", "Caption Change", _
 Left$(WndCaption, R))

' make sure caption is not empty
If ChangeCaption = "" Then
 MsgBox "Thanks anyway", vbInformation, "API Spyer"
Else
  ' confirm change
  If MsgBox("Change caption to ''" & ChangeCaption & "''?", vbQuestion + vbYesNo, _
   "Confirm Change") = vbYes Then SetWindowText WinHwnd, ChangeCaption
End If
End Sub




Private Sub mnuChangeEnabled_Click()
mnuChangeEnabled.Checked = Not (mnuChangeEnabled.Checked)
If mnuChangeEnabled.Checked Then
  EnableWindow WinHwnd, True
Else
  EnableWindow WinHwnd, False
End If
End Sub

Private Sub mnuChangeVisible_Click()
mnuChangeVisible.Checked = Not (mnuChangeVisible.Checked)
If mnuChangeVisible.Checked Then
  ShowWindow WinHwnd, True
Else
  ShowWindow WinHwnd, False
End If
End Sub


Private Sub mnuGethDC_Click()
MsgBox "hWnd: " & CStr(WinHwnd) & vbCrLf & "hDC: " & CStr(GetDC(WinHwnd)), vbInformation, "API Spyer"
End Sub

Private Sub mnuGetRect_Click()
' get the Rect of the window
Dim Rct As RECT
GetWindowRect WinHwnd, Rct
' display info
MsgBox "hWnd: " & CStr(WinHwnd) & vbCrLf & "Left: " & CStr(Rct.Left) & vbCrLf & "Top: " & _
 CStr(Rct.Top) & vbCrLf & "Bottom: " & CStr(Rct.Bottom) & vbCrLf & "Right: " & CStr(Rct.Right)
End Sub


Private Sub mnuHelpAbout_Click()
MsgBox "API Spyer" & vbCrLf & "Written by Steve Weller", vbInformation, "About API Spyer..."
End Sub

Private Sub mnuHelpTopics_Click()
' show message box help
MsgBox "Just make sure that the API Spyer window has the focus.  Point the mouse over the window you want information on.  If you want to change (or get) various values, hit Enter to go to Change Window mode.  The Items are" & vbCrLf & _
 "Change:" & vbCrLf & _
 "  Caption:  Changes the window's caption" & vbCrLf & _
 "  Visible:  Whether the window is visible" & vbCrLf & _
 "Get:" & vbCrLf & _
 "  hDC:  The window's display context handle" & vbCrLf & _
 "  Window Rect:  The window's position (in pixels)" & vbCrLf & _
 "Misc:" & vbCrLf & _
 "  Bring To Top:  Brings the window to the top" & vbCrLf & _
 "Help:" & vbCrLf & _
 "  Topics:  This dialog box" & vbCrLf & _
 "  About:  An About box showing the program's writer"
End Sub

Private Sub mnuMiscBringToTop_Click()
' bring the window to the top
BringWindowToTop WinHwnd
End Sub


