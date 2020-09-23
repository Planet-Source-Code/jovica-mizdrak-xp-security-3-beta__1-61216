Attribute VB_Name = "mMain"
Option Explicit



Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Const GWL_EXSTYLE = -20
    Public Const WS_EX_LAYERED = &H80000
    Public Const GWL_STYLE = (-16)
    Public Const WS_VISIBLE = &H10000000
    
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Public Const LWA_ALPHA = &H2
    
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Const DESKTOP_NAME As String = "XpSecurity"
Public strPassword As String, Global_Lock As Boolean
Public strFRM As String, strOpt As String, strArg(255) As String

Private Function AppPath(ByVal zPath As String) As String
  If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"
End Function

Private Function FileExist(ByVal strPath As String) As Boolean
  On Local Error GoTo ErrFile
  Open strPath For Input Access Read As #1
  Close #1
  FileExist = True
  Exit Function
ErrFile:
  FileExist = False
End Function

Private Sub MakeManifest()
  Dim file$, file2$, qwe As String
  file$ = AppPath(App.Path) & App.EXEName & ".exe.MANIFEST"
  If Not FileExist(file$) Then
    qwe = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf _
        & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf _
        & "<assemblyIdentity type=""win32"" processorArchitecture=""*"" version=""6.0.0.0"" name=""name""/>" & vbCrLf _
        & "<description>Enter your Description Here</description>" & vbCrLf _
        & "<dependency>" & vbCrLf _
        & "   <dependentAssembly>" & vbCrLf _
        & "      <assemblyIdentity" & vbCrLf _
        & "           type=""win32""" & vbCrLf _
        & "           name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0""" & vbCrLf _
        & "           language=""*""" & vbCrLf _
        & "           processorArchitecture=""*""" & vbCrLf _
        & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf _
        & "      />" & vbCrLf _
        & "   </dependentAssembly>" & vbCrLf _
        & "</dependency>" & vbCrLf _
        & "</assembly>" & vbCrLf
    Open file$ For Binary Access Write Lock Write As #1 Len = 1
    Put #1, , qwe
    Close #1
    SetAttr file$, vbReadOnly Or vbHidden 'Or vbSystem
    file2$ = AppPath(App.Path) & App.EXEName & ".exe"
    Shell file2$, vbNormalFocus
    ExitProcess 1
  End If
End Sub

Public Sub InitControlsXP()
  MakeManifest
  InitCommonControls
End Sub

Public Function GetArgs(ArgSep As String) As Collection
    Dim Arr     As Variant
    Dim i       As Integer
    Dim Arg     As String
    Dim Col     As New Collection
    
    Arg = Command$
    Arg = Right(Arg, Len(Arg) - 1)
        
    Arr = Split(Arg, ArgSep, , vbTextCompare)
    
    For i = 0 To UBound(Arr)
        Col.Add Arr(i)
    Next i
    Set GetArgs = Col
    
    Set Col = Nothing
End Function
