VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer FadeOUT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer FadeIN 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ifade As Integer, iCount As Integer

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

Private Sub Timer1_Timer()
    FadeOUT.Enabled = True
    Timer1.Enabled = False
End Sub

