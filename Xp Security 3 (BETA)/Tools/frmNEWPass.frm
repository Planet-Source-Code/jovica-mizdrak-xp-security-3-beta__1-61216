VERSION 5.00
Begin VB.Form frmNEWPass 
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNEWPass.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer FadeIN 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer FadeOUT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.TextBox txtPass 
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
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1080
      Width           =   3345
   End
   Begin VB.TextBox txtPass 
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
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2420
      Width           =   3345
   End
   Begin VB.TextBox txtPass 
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
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1820
      Width           =   3345
   End
   Begin VB.Label lblCheck 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password Mismatch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCheck 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   795
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   365
      Left            =   4280
      Top             =   2790
      Width           =   1260
   End
   Begin VB.Image img_ok 
      Height          =   355
      Left            =   3370
      Top             =   2790
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CANCEL"
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
      Left            =   4600
      TabIndex        =   1
      Top             =   2870
      Width           =   735
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
      Left            =   3665
      TabIndex        =   0
      Top             =   2870
      Width           =   255
   End
End
Attribute VB_Name = "frmNEWPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ifade As Integer, iCount As Integer, iTmr As Integer, iChk As Integer

Private Sub Form_Load()
    GLOB_Move = False
    iCount = 0
    ifade = 255
    SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, ifade, LWA_ALPHA
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub FadeIN_Timer()
    If ifade = 255 Then
        iCount = iCount + 1
        FadeIN.Enabled = False
        FadeOUT.Enabled = True
    Else
        ifade = ifade + 15
        SetLayeredWindowAttributes Me.hWnd, 0, ifade, LWA_ALPHA
    End If
End Sub

Private Sub FadeOUT_Timer()
    If ifade <= 120 Then
        If iCount = 3 Then
            ifade = ifade - 15
            SetLayeredWindowAttributes Me.hWnd, 0, ifade, LWA_ALPHA
            If ifade = 0 Then
                Unload Me
                Form1.Show
                GLOB_Move = True
            End If
        Else
            FadeOUT.Enabled = False
            FadeIN.Enabled = True
        End If
    
    Else
        ifade = ifade - 15
        SetLayeredWindowAttributes Me.hWnd, 0, ifade, LWA_ALPHA
    End If
End Sub

Private Sub Image1_Click()
Timer1.Enabled = True
End Sub

Private Sub img_ok_Click()
If txtPass(0).Text = strPassword Then
    If txtPass(1).Text = txtPass(2).Text Then
        strPassword = txtPass(1).Text
        SaveSetting App.EXEName, "p1", 1, strPassword
        Timer1.Enabled = True
    Else
        iChk = 1
        Timer2.Enabled = True
    End If
Else
iChk = 0
Timer2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
    FadeOUT.Enabled = True
    Me.Enabled = False
    Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()
iTmr = iTmr + 1
If iTmr = 50 And iChk = 0 Then
lblCheck(0).Visible = False
iTmr = 0
Timer2.Enabled = False
Else
lblCheck(0).Visible = True
End If
If iTmr = 50 And iChk = 1 Then
lblCheck(1).Visible = False
Else
lblCheck(1).Visible = True
iTmr = 0
Timer2.Enabled = False
End If
End Sub
