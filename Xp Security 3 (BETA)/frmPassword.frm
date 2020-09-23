VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Configuration - Xp Security"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2370
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1410
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   870
      Width           =   4695
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-enter New Password:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Picture         =   "frmPassword.frx":51F0
      Top             =   2280
      Width           =   4680
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
cmdApply.Enabled = False
If txtPass(0).Text = strPassword Then
    If txtPass(1).Text = txtPass(2).Text Then
        strPassword = txtPass(1).Text
        cmdApply.Enabled = False
    Else
    MsgBox "New passwords don't match." & vbNewLine & "Please try again!", vbCritical + vbOKOnly, "Xp Security"
    Call BlankTxt
    End If
Else
MsgBox "Old password is incorrect." & vbNewLine & "Please try again!", vbCritical + vbOKOnly, "Xp Security"
Call BlankTxt
End If
End Sub

Function BlankTxt()
txtPass(0).Text = ""
txtPass(1).Text = ""
txtPass(2).Text = ""
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If cmdApply.Enabled Then cmdApply_Click
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form1.Visible = True Then
    Unload Me
Else
    Global_Lock = True
    TrayChangeIcon Form1, App.Path & "\Res\1.ico", "XP Security"
End If
End Sub

Private Sub txtPass_Change(Index As Integer)
cmdApply.Enabled = True
End Sub
