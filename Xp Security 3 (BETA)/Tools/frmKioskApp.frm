VERSION 5.00
Begin VB.Form frmKioskApp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1425
   ClientLeft      =   4050
   ClientTop       =   4530
   ClientWidth     =   2490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKioskApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   480
   End
End
Attribute VB_Name = "frmKioskApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cTile As New cDIBTile
Private cTile2 As New cDIBTile
Private cTop As New cDIBTile
Private cBot As New cDIBTile
Public hhkLowLevelKybd As Long

Private Sub Form_Load()
        '-- Set pattern
    cTile.SetPattern LoadResPicture("103", vbResBitmap)
    cTile2.SetPattern LoadResPicture("104", vbResBitmap)
    cTop.SetPattern LoadResPicture("102", vbResBitmap)
    cBot.SetPattern LoadResPicture("101", vbResBitmap)
    SetWindowPos Me.hWnd, 1, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    strPassword = GetSetting(App.EXEName, "p1", 1, "jovica")
End Sub

Private Sub Command6_Click()
'On Error GoTo err:
'Static bolTaskMgr As Boolean
'If bolTaskMgr = False Then
'SetAttr "C:\windows\system32\taskmgr.exe", vbHidden + vbSystem
'Open "C:\windows\system32\taskmgr.exe" For Binary As #1
'bolTaskMgr = True
'Else
'Close #1
'bolTaskMgr = False
'End If
'err: Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GLOB_Move = True Then
        Form1.FadeIN.Enabled = False
        Form1.FadeOUT.Enabled = True
    End If
End Sub

Private Sub Form_Paint()
    '-- Tile pattern
    cTile.Tile hdc, 0, 0, ScaleWidth, 197
    cTile2.Tile hdc, 0, ScaleHeight - 197, ScaleWidth, ScaleHeight
    cTop.Tile hdc, ScaleWidth / 2 - (413 / 2), 0, ScaleWidth / 2 + (413 / 2), 129
    cBot.Tile hdc, ScaleWidth / 2 - (413 / 2), ScaleHeight - 129, ScaleWidth / 2 + (413 / 2), ScaleHeight
End Sub

Private Sub Form_Resize()
    hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
    'Command6_Click
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "p1", 1, strPassword
    UnhookWindowsHookEx hhkLowLevelKybd
    Set cTile = Nothing
    Set cTile2 = Nothing
    Set cTop = Nothing
    Set cBot = Nothing
    Set frmKioskApp = Nothing
End Sub

Private Sub Timer1_Timer()
If GLOB_Move = True Then Me.SetFocus
End Sub
