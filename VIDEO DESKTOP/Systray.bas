Attribute VB_Name = "Systray"
Option Explicit
Type NOTIFYICONDATA
    cbSize                        As Long
    hwnd                          As Long
    uId                           As Long
    uFlags                        As Long
    ucallbackMessage              As Long
    hIcon                         As Long
    szTip                         As String * 64
End Type
Private Const NIM_ADD         As Long = &H0
Private Const NIM_DELETE      As Long = &H2
Private Const NIF_MESSAGE     As Long = &H1
Private Const NIF_ICON        As Long = &H2
Private Const NIF_TIP         As Long = &H4
Private Const WM_MOUSEMOVE    As Long = &H200
Public Const WM_RBUTTONUP     As Long = &H205
Private iData                 As NOTIFYICONDATA
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          ByVal lpsz2 As String) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                       lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Sub RemoveIconTray(DataIcon As PictureBox, _
                          ByVal zTip As String)
    With iData
        .cbSize = Len(iData)
        .hwnd = Form1.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = DataIcon
        .szTip = zTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_DELETE, iData
End Sub
Public Sub SetIconTray(DataIcon As PictureBox, _
                       ByVal zTip As String)
    With iData
        .cbSize = Len(iData)
        .hwnd = Form1.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = DataIcon
        .szTip = zTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, iData
End Sub


