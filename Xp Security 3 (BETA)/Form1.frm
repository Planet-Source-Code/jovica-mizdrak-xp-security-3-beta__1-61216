VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xp Security 3 (BETA)"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "Options"
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Passwords"
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Task Manager"
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   34
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Firewall"
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   33
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Security Center"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   32
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Automatic Updates"
      Height          =   255
      Index           =   12
      Left            =   3960
      TabIndex        =   31
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Add or Remove Programs"
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   30
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Timer tmrDisablePA 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lock Windows"
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Disable Task Manager"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "More Security"
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "User Accounts 2"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   26
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "User Accounts"
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   25
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Taskbar Properties"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   24
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Internet Properties"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   23
      Top             =   4200
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Folder Options"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   22
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Display Properties"
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   21
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "System Properties"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   20
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Disable Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deactivate"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   18
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame frameDPA 
      Caption         =   "Disable Programm Access"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Administrative Tools"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   15
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "CMD Prompt"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Control Panel"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Registry Editor"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      Begin VB.CheckBox chkDisable 
         Caption         =   "WIN + R"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   35
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "SHIFT + ENTER"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   11
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "APP POPUP"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   10
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "WIN + L"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   9
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "CTRL + ESCAPE"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   8
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + F4"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + ENTER"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + SPACE"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + TAB"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + ESCAPE"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1080
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   0
      Width           =   6195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Menu mSysPopup 
      Caption         =   "SysPopup"
      Visible         =   0   'False
      Begin VB.Menu mShow 
         Caption         =   "Open Xp Security"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mED 
         Caption         =   "Deactivate"
      End
      Begin VB.Menu mOption 
         Caption         =   "Xp Security Options"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'this API is necessary to make sure that menu will disappear if user clicks outside of it
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim hhkLowLevelKybd As Long, iExit As Boolean, bolTaskMgr As Boolean

Private Sub Command2_Click()
Global_Lock = True
Form_Unload CInt(iExit)
End Sub

Private Sub Command3_Click()
frmPassword.Show vbModal, Me
End Sub

Private Sub Command4_Click()
On Error GoTo err:
   Dim sPath As String
   Dim cNewDesktop As New cDesktop
   cNewDesktop.Create DESKTOP_NAME
   sPath = App.Path & "\ldesk.exe"
   cNewDesktop.StartProcess sPath
   Call Command2_Click
Exit Sub
err:
MsgBox "Error number: " & err.Number & vbNewLine & "Description: " & err.Description, vbCritical
End Sub

Private Sub Command7_Click()
frmOptions.Show vbModal, Me
End Sub

Private Sub Form_Initialize()
    InitControlsXP
End Sub

Public Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        Command1(0).Enabled = False
        Command1(1).Enabled = True
        mED.Caption = "Deactivate"
        strFRM = "fmain"
        tmrDisablePA.Enabled = True
    Case 1
        UnhookWindowsHookEx hhkLowLevelKybd
        hhkLowLevelKybd = 0
        Command1(0).Enabled = True
        Command1(1).Enabled = False
        mED.Caption = "Activate"
        tmrDisablePA.Enabled = False
End Select
End Sub

Private Sub Form_Load()
'
'On instance<<<<<<<<<<<<<
'

On Error GoTo err:
    Dim Args As Collection
    SetIcon Me.hwnd, "AA0", True
    iExit = False
    strPassword = GetSetting(App.EXEName, "p1", 1, "jovica")
    strOpt = GetSetting(App.EXEName, "o1", 2, "none")
    Global_Lock = True
    TrayAddIcon Form1, App.Path & "\Res\1.ico", "XP Security"
    Command1_Click 0
    If Command$ <> "" Then
        Set Args = GetArgs(" -")
        Dim i%
        For i = 1 To Args.Count
            strArg(i) = Args(i)
        Next i
        If strArg(1) = "lockwin" Or strArg(1) = "lockwina" Then Call Command4_Click
    Else
        strArg(1) = ""
    End If
Exit Sub
err:
MsgBox "Form Load" & vbNewLine & "Error number: " & err.Number & vbNewLine & "Description: " & err.Description, vbCritical
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    'All mouse events including balloon click
    Dim Result As Long
    Dim cEventx As Single
    Dim cEventy As Single
    cEventx = x / Screen.TwipsPerPixelX
    cEventy = Y / Screen.TwipsPerPixelY

Select Case cEventx Xor cEventy
    Case MouseMove
        'Debug.Print "MouseMove"
    Case LeftUp
        'Debug.Print "Left Up"
    Case LeftDown
        'Debug.Print "LeftDown"
    Case LeftDbClick
        'Debug.Print "LeftDbClick"
        If Global_Lock = True Then
        strFRM = "fmain"
        frmPassPrompt.Show
        Else
        Me.WindowState = 0
        Me.Show
        End If
    Case MiddleUp
        'Debug.Print "MiddleUp"
    Case MiddleDown
        'Debug.Print "MiddleDown"
    Case MiddleDbClick
        'Debug.Print "MiddleDbClick"
    Case RightUp
        'Debug.Print "RightUp"
        'now show it
        PopupMenu mSysPopup, , , , mShow
    Case RightDown
        'Debug.Print "RightDown"
        'make sure that menu will disappear if user clicks outside of it
        Result = SetForegroundWindow(Me.hwnd)
    Case RightDbClick
        'Debug.Print "RightDbClick"
    Case BalloonClick
        'Debug.Print "Balloon Click"

    End Select
End Sub

Function changeTrayIcon()
If Global_Lock = True Then
TrayChangeIcon Form1, App.Path & "\Res\1.ico", "XP Security"
ElseIf Global_Lock = False Then
TrayChangeIcon Form1, App.Path & "\Res\2.ico", "XP Security"
End If
End Function

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
        Me.Hide
        Call changeTrayIcon
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = iExit
If Cancel = True Then
    TrayRemoveIcon
    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
Else
    Cancel = True
    Me.Hide
    Call changeTrayIcon
End If
End Sub

Private Sub mED_Click()
If mED.Caption = "Deactivate" Then
If Not Form1.Visible Then
    Global_Lock = True
    changeTrayIcon
End If
If Global_Lock = False Then
Command1_Click 1
Else
strFRM = "Deactivate"
frmPassPrompt.Show vbModal, Me
End If
Global_Lock = False
ElseIf mED.Caption = "Activate" Then
mED.Caption = "Deactivate"
If Not Form1.Visible Then
    Global_Lock = True
    changeTrayIcon
End If
Command1_Click 0
Global_Lock = True
If Form1.Visible Then
    Global_Lock = False
    'changeTrayIcon
End If
End If
End Sub

Private Sub mOption_Click()
On Error Resume Next
If Not Global_Lock = True Then
    frmOptions.Show vbModal, Me
Else
    strFRM = "fopt"
    frmPassPrompt.Show vbModal, Me
End If
End Sub

Private Sub mShow_Click()
    If Global_Lock = True Then
        strFRM = "fmain"
        frmPassPrompt.Show
    Else
        Me.WindowState = 0
        Me.Show
    End If
End Sub

Private Sub mExit_Click()
    iExit = True
    Unload Me
    End
End Sub

Private Sub tmrDisablePA_Timer()
On Error Resume Next
Dim WindowToFind As Long    'Window Handle
Dim ChildWin As Long
Dim ParentWin As Long
Dim RunDll As String
RunDll = "C:\WINDOWS\system32\rundll32.exe"
If chkDisablePA(0).Value = 1 Then
    WindowToFind& = FindWindow("RegEdit_RegEdit", "Registry Editor") ' Look for "Registry Editor"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(1).Value = 1 Then
WindowToFind& = FindWindow("CabinetWClass", "Control Panel") ' Look for "Control Panel"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(2).Value = 1 Then
    WindowToFind& = FindWindow("ConsoleWindowClass", "C:\WINDOWS\system32\cmd.exe") ' Look for "CMD"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(3).Value = 1 Then ' Administrative Tools
    WindowToFind& = FindWindow("CabinetWClass", "Administrative Tools")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(4).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) 'System Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "System Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(5).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) ' Display Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Display Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(6).Value = 1 Then 'Folder Options
ParentWin& = FindWindow("MSGlobalFolderOptionsStub", "Folder Options")
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Folder Options")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(7).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) 'Internet Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Internet Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(8).Value = 1 Then 'Taskbar and Start Menu Properties
ParentWin& = FindWindow("Static", "Taskbar and Start Menu Properties")
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Taskbar and Start Menu Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(9).Value = 1 Then 'User Accounts
    WindowToFind& = FindWindow("HTML Application Host Window Class", "User Accounts")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(10).Value = 1 Then 'User Accounts2
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "User Accounts")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(11).Value = 1 Then '"Add or Remove Programs"
WindowToFind& = FindWindow("NativeHWNDHost", "Add or Remove Programs")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(12).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Automatic Updates")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(13).Value = 1 Then
WindowToFind& = FindWindow("wscui_class", "Windows Security Center")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If chkDisablePA(14).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Windows Firewall")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If


If chkDisablePA(15).Value = 1 Then
WindowToFind& = FindWindow("#32770", "Windows Task Manager")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If
End Sub
