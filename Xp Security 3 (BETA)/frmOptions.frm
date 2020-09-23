VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xp Security - Options"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOptions.frx":0000
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   0
      ScaleHeight     =   2250
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   870
      Width           =   4695
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dont Run Xp Security at Start Up"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Xp Security at Start Up"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   4455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Xp Security at Start Up and Lock Windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   4455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Xp Security at Start Up and Lock Windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   4455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Xp Security at Start Up"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblCAption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "This setting applys to All users:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblCAption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "This setting applys to Current user:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Picture         =   "frmOptions.frx":509D
      Top             =   3120
      Width           =   4680
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If opt(0).Value = True Then
        CreateKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title, IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe"
        strOpt = "nolockwin"
    Else
        DeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
    End If
    
    If opt(1).Value = True Then
        CreateKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title, """" + App.Path + "\" + App.EXEName + ".exe" + """ -lockwin"
        strOpt = "lockwin"
    Else
        DeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
    End If

    If opt(2).Value = True Then
        CreateKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title, IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe"
        strOpt = "nolockwina"
    Else
        DeleteKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
    End If
    
    If opt(3).Value = True Then
        CreateKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title, """" + App.Path + "\" + App.EXEName + ".exe" + """ -lockwina"
        strOpt = "lockwina"
    Else
        DeleteKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
    End If

    If opt(4).Value = True Then
        DeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
        DeleteKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" & App.Title
        strOpt = "none"
    End If
    
    SaveSetting App.EXEName, "o1", 2, strOpt
    cmdOk.SetFocus
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If cmdApply.Enabled Then cmdApply_Click
Unload Me
End Sub

Private Sub Form_Load()
Call LoadSettings
End Sub

Private Sub Form_Resize()
Call LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form1.Visible = True Then
    Unload Me
Else
    Global_Lock = True
    TrayChangeIcon Form1, App.Path & "\Res\1.ico", "XP Security"
End If
End Sub

Function LoadSettings()
If strOpt = "nolockwin" Then
opt(0).Value = True
ElseIf strOpt = "lockwin" Then
opt(1).Value = True
ElseIf strOpt = "nolockwina" Then
opt(2).Value = True
ElseIf strOpt = "lockwina" Then
opt(3).Value = True
ElseIf strOpt = "none" Then
opt(4).Value = True
End If

cmdApply.Enabled = False
End Function

Private Sub opt_Click(Index As Integer)
cmdApply.Enabled = True
End Sub
