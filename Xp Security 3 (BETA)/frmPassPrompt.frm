VERSION 5.00
Begin VB.Form frmPassPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password - Xp Security"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdBTN 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBTN 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Picture         =   "frmPassPrompt.frx":0000
      Top             =   0
      Width           =   4680
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBTN_Click(Index As Integer)
Select Case Index
Case 0:
        If txtPass.Text = strPassword Then
            Global_Lock = False
            TrayChangeIcon Form1, App.Path & "\Res\2.ico", "XP Security"
            If strFRM = "fmain" Then
                Form1.WindowState = 0
                Form1.Show
            ElseIf strFRM = "fopt" Then
                frmOptions.Show vbModal, Form1
            ElseIf strFRM = "Deactivate" Then
                Form1.mED.Caption = "Activate"
                Form1.Command1_Click 1
            End If
            Unload Me
        Else
            Unload Me
            MsgBox "Incorrect password!", vbCritical + vbOKOnly
        End If
Case 1: Unload Me
End Select
End Sub

Private Sub Form_Resize()
txtPass.SetFocus
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtPass.Enabled = False
        txtPass.Locked = True
        Call cmdBTN_Click(0)
    End If
End Sub
