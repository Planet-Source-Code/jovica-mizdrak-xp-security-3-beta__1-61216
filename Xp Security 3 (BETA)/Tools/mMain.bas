Attribute VB_Name = "mMain"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Const GWL_EXSTYLE = -20
    Public Const WS_EX_LAYERED = &H80000
    Public Const GWL_STYLE = (-16)
    Public Const WS_VISIBLE = &H10000000
    
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Const LWA_ALPHA = &H2
    
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function OpenInputDesktop Lib "user32" ( _
      ByVal dwFlags As Long, _
      ByVal fInherit As Boolean, _
      ByVal dwDesiredAccess As Long _
   ) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationW" ( _
   ByVal hObj As Long, _
   ByVal nIndex As Long, _
   pvInfo As Any, _
   ByVal nLength As Long, _
   lpnLengthNeeded As Long) As Long
Private Const UOI_FLAGS = 1
Private Const UOI_NAME = 2
Private Const UOI_TYPE = 3
Private Const UOI_USER_SID = 4
Private Const DESKTOP_READOBJECTS = &H1&
Private Const DESKTOP_NAME As String = "XpSecurity"
Public intFade As Integer, GLOB_Move As Boolean, strPassword As String

Public Sub Main()
   InitCommonControls
    intFade = 255
   ' Check we are running in the correct desktop
   If (GetDesktopName() = DESKTOP_NAME) Then
      
      ' Ok, let's run
      Dim fK As New frmKioskApp
      fK.Show
   Else
   
      MsgBox "This application cannot be run directly.", vbCritical
   
   End If
   
End Sub

Public Function GetDesktopName() As String
Dim hDesktop As Long
Dim lR As Long
Dim lSize As Long
Dim sBuff As String
Dim iPos As Long
   
   hDesktop = OpenInputDesktop(0, False, DESKTOP_READOBJECTS)
   If Not (hDesktop = 0) Then
      lSize = (Len(DESKTOP_NAME) + 1) * 2
      ReDim bBuff(0 To lSize - 1) As Byte
      lR = GetUserObjectInformation(hDesktop, UOI_NAME, bBuff(0), lSize, lSize)
      sBuff = bBuff
      iPos = InStr(sBuff, vbNullChar)
      If (iPos > 1) Then
         sBuff = Left(sBuff, iPos - 1)
      End If
      GetDesktopName = sBuff
      CloseHandle hDesktop
   End If
End Function
