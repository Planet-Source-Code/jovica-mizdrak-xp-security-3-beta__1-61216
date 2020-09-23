VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1100
      Width           =   2505
   End
   Begin VB.Timer FadeIN 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer FadeOUT 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image img_ok 
      Height          =   375
      Left            =   4600
      Top             =   1000
      Width           =   815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4890
      TabIndex        =   1
      Top             =   1090
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Text1.Visible = True
    Text1.Locked = False
    Text1.SetFocus
    GLOB_Move = False
End Sub

Private Sub Form_Load()
    GLOB_Move = True
    SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub FadeIN_Timer()
    If intFade = 255 Then
        FadeIN.Enabled = False
    Else
        intFade = intFade + 5
        SetLayeredWindowAttributes Me.hWnd, 0, intFade, LWA_ALPHA
    End If
End Sub

Private Sub FadeOUT_Timer()
    If intFade = "80" Then
        FadeOUT = False
    Else
        intFade = intFade - 5
        SetLayeredWindowAttributes Me.hWnd, 0, intFade, LWA_ALPHA
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GLOB_Move = True Then
        FadeOUT.Enabled = False
        FadeIN.Enabled = True
    End If
End Sub

Private Sub img_ok_Click()
    If Text1.Text = strPassword Then
            Do Until intFade <= 5
                intFade = intFade - 5
                If intFade <= 5 Then Exit Do
                SetLayeredWindowAttributes Me.hWnd, 0, intFade, LWA_ALPHA
            Loop
            Unload Me
            Form2.Show
            GLOB_Move = False
    Else
        Text1.Text = ""
        Text1.Enabled = True
        Unload Me
        Form3.Show
    End If
End Sub

Private Sub Label2_Click()
            Do Until intFade <= 5
                intFade = intFade - 5
                If intFade <= 5 Then Exit Do
                SetLayeredWindowAttributes Me.hWnd, 0, intFade, LWA_ALPHA
            Loop
            Unload Me
            frmNEWPass.Show
            GLOB_Move = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.FontBold = True
End Sub

Private Sub Text1_GotFocus()
    GLOB_Move = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Text1.Enabled = False
        Text1.Locked = True
        Call img_ok_Click
    ElseIf KeyCode = 27 Then
        Text1.Locked = True
        Text1.Text = ""
        GLOB_Move = True
    End If
End Sub

Private Sub Text1_LostFocus()
    GLOB_Move = True
End Sub
