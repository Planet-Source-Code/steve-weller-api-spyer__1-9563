VERSION 5.00
Begin VB.Form frmAPISpy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Spyer"
   ClientHeight    =   2715
   ClientLeft      =   2415
   ClientTop       =   2265
   ClientWidth     =   3135
   Icon            =   "ApiSpy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3135
   Begin VB.CommandButton cmdChangeWindow 
      Caption         =   "&Change Window"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "&On Top"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Timer tmrInfo 
      Interval        =   150
      Left            =   2640
      Top             =   120
   End
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmAPISpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2
Private Function GetInformation(ParamArray HwndExcluded() As Variant) As String
' get information about
' window mouse is over
On Error Resume Next

Dim CursorPos As POINTAPI
Dim szText As String * 100
Dim HoldText As String
Dim HwndNow As Long, hInst As Long
Dim Rct As RECT, R As Long
Dim I
Static HwndPrev As Long

Const GWW_HINSTANCE = (-6), GWW_ID = (-12), GWL_STYLE = (-16)

GetCursorPos CursorPos

HwndNow = WindowFromPoint(CursorPos.x, CursorPos.y)

For I = LBound(HwndExcluded) To UBound(HwndExcluded)
  If HwndNow = CLng(HwndExcluded(I)) Then Exit Function
Next I

If HwndNow <> HwndPrev Then
  HwndPrev = HwndNow
  
  frmAPIChange.WinHwnd = HwndNow
  
  HoldText = HoldText & "Handle of window is: &H" & Hex$(HwndNow) & vbCrLf
  
  R = GetWindowText(HwndNow, szText, 100)
  HoldText = HoldText & "Title of window is: " & Left(szText, R) & vbCrLf
  
  R = GetClassName(HwndNow, szText, 100)
  HoldText = HoldText & "It's class name is: " & Left(szText, R) & vbCrLf
  
  GetWindowRect HwndNow, Rct
  HoldText = HoldText & "It's width is: " & CStr(Rct.Right - Rct.Left) & vbCrLf
  HoldText = HoldText & "It's height is: " & CStr(Rct.Bottom - Rct.Top) & vbCrLf
    
  If IsIconic(HwndNow) Then
    HoldText = HoldText & "Window is Minimized" & vbCrLf
  ElseIf IsZoomed(HwndNow) Then
    HoldText = HoldText & "Window is Maximized" & vbCrLf
  Else
    HoldText = HoldText & "Window is Normal" & vbCrLf
  End If
  
  HoldText = HoldText & "Keyboard Input " & IIf(IsWindowEnabled(HwndNow), "", "NOT ") & "available" & vbCrLf
  
  HoldText = HoldText & "Application Instance is: " & GetWindowWord(HwndNow, GWW_HINSTANCE) & vbCrLf
      
  GetInformation = HoldText
End If
End Function





Private Sub chkOnTop_Click()
' change whether window is on top or not
If chkOnTop.Value = vbChecked Then
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Else
  SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Private Sub cmdChangeWindow_Click()
' show the Change Window form
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
tmrInfo.Enabled = False
frmAPIChange.Show vbModal
End Sub

Private Sub Form_Load()
' search for first instance
If App.PrevInstance Then End

' set window as topmost
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
tmrInfo.Enabled = False
If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "APISpyer by Steve Weller") = vbNo Then
  tmrInfo.Enabled = True
  Cancel = True
End If
End Sub

Private Sub tmrInfo_Timer()
' get information from window
' that mouse is over
On Error Resume Next

' yield events for program to OS
DoEvents

Dim Info As String
Info = GetInformation(Me.hwnd, txtInfo.hwnd, chkOnTop.hwnd, cmdChangeWindow.hwnd)

If Info <> "" Then
  txtInfo = Info
End If
End Sub


