Attribute VB_Name = "mWinFun"
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Sub Hide_DeskTop()
Dim progman As Long
Dim shelldlldefview As Long
Dim SysListView As Long
progman = FindWindow("progman", vbNullString)
shelldlldefview = FindWindowEx(progman, 0&, "shelldll_defview", vbNullString)
SysListView = FindWindowEx(shelldlldefview, 0&, "syslistview32", vbNullString)
Call ShowWindow(SysListView, SW_HIDE)
End Sub
Sub Show_DeskTop()
Dim progman As Long
Dim shelldlldefview As Long
Dim SysListView As Long
progman = FindWindow("progman", vbNullString)
shelldlldefview = FindWindowEx(progman, 0&, "shelldll_defview", vbNullString)
SysListView = FindWindowEx(shelldlldefview, 0&, "syslistview32", vbNullString)
Call ShowWindow(SysListView, SW_SHOW)
End Sub

Sub Hide_Clock()
Dim ShellTrayWnd As Long
Dim TrayNotifyWnd As Long
Dim TrayClockWClass As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "traynotifywnd", vbNullString)
TrayClockWClass = FindWindowEx(TrayNotifyWnd, 0&, "trayclockwclass", vbNullString)
Call ShowWindow(TrayClockWClass, SW_HIDE)
End Sub
Sub Show_Clock()
Dim ShellTrayWnd As Long
Dim TrayNotifyWnd As Long
Dim TrayClockWClass As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "traynotifywnd", vbNullString)
TrayClockWClass = FindWindowEx(TrayNotifyWnd, 0&, "trayclockwclass", vbNullString)
Call ShowWindow(TrayClockWClass, SW_SHOW)
End Sub
Sub Hide_Start()
Dim ShellTrayWnd As Long
Dim button As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(ShellTrayWnd, 0&, "button", vbNullString)
Call ShowWindow(button, SW_HIDE)
End Sub
Sub Show_Start()
Dim ShellTrayWnd As Long
Dim button As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(ShellTrayWnd, 0&, "button", vbNullString)
Call ShowWindow(button, SW_SHOW)
End Sub
Sub Hide_TaskBar()
Dim ShellTrayWnd As Long
Dim ReBarWindow As Long
Dim MSTaskSwWClass As Long
Dim ToolbarWindow As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
ReBarWindow = FindWindowEx(ShellTrayWnd, 0&, "rebarwindow32", vbNullString)
MSTaskSwWClass = FindWindowEx(ReBarWindow, 0&, "mstaskswwclass", vbNullString)
ToolbarWindow = FindWindowEx(MSTaskSwWClass, 0&, "toolbarwindow32", vbNullString)
Call ShowWindow(ToolbarWindow, SW_HIDE)
End Sub
Sub Show_TaskBar()
Dim ShellTrayWnd As Long
Dim ReBarWindow As Long
Dim MSTaskSwWClass As Long
Dim ToolbarWindow As Long
ShellTrayWnd = FindWindow("shell_traywnd", vbNullString)
ReBarWindow = FindWindowEx(ShellTrayWnd, 0&, "rebarwindow32", vbNullString)
MSTaskSwWClass = FindWindowEx(ReBarWindow, 0&, "mstaskswwclass", vbNullString)
ToolbarWindow = FindWindowEx(MSTaskSwWClass, 0&, "toolbarwindow32", vbNullString)
Call ShowWindow(ToolbarWindow, SW_SHOW)
End Sub
