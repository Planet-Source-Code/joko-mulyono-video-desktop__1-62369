VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   0
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4710
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrLoop 
      Left            =   3000
      Top             =   2640
   End
   Begin VB.PictureBox imgIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5640
      Picture         =   "Form1.frx":34CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Tag             =   "pic"
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   1440
      Picture         =   "Form1.frx":6994
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Option"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuFName 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDT 
         Caption         =   "Tool"
         Begin VB.Menu mnuHD 
            Caption         =   "Hide Desktop"
            Shortcut        =   +{DEL}
         End
         Begin VB.Menu mnuSD 
            Caption         =   "Show Desktop"
            Enabled         =   0   'False
            Shortcut        =   +{INSERT}
         End
         Begin VB.Menu mnuloop 
            Caption         =   "Loop"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSt 
            Caption         =   "StartUp"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCon 
         Caption         =   "Control"
         Enabled         =   0   'False
      End
      Begin VB.Menu spc3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuFF 
         Caption         =   "FForward"
      End
      Begin VB.Menu mnuFR 
         Caption         =   "FRewind"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuM 
         Caption         =   "Mute"
      End
      Begin VB.Menu spc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
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

Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const CCM_FIRST             As Long = &H2000
Private Const CCM_SETCOLORSCHEME    As Long = (CCM_FIRST + 2)
Private Const GWL_WNDPROC           As Long = -4
Private Const MOD_CONTROL           As Long = &H2
Private MM                          As Object
Private isPlaying                   As Boolean
Private isPaused                    As Boolean
Private isMute                      As Boolean
Private filmname                    As String
Private lngMsg                      As Long
Private blnFlag                     As Boolean
Private StarUP                      As Boolean
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Any, _
                                                      ByVal bErase As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long

Private Sub CLOSEALLPLAYER()

    With MM
        .setCommand StopCD
        .setCommand CloseCD
    End With
    WALLPAPERTHEME (RestoreWall)
    filmname = ""
    tmrLoop.Interval = 0
End Sub
Private Sub Form_Initialize()
    If App.PrevInstance Then
        End
    End If
    isPaused = False
    isMute = False
    ScreenSaverActive False
End Sub
Private Sub Form_Load()

    On Error Resume Next
    Set MM = New VIDEODESKTOP
    MM.HwndParent = FindWindow(vbNullString, "Program Manager")
    startODMenus Me, True
    With CustomMenu
        .Texture = True
        Set .Picture = Image1.Picture
        .UseCustomFonts = False
'        .FontUnderline = False
'        .FontName = "Lucida Sans"
'        .FontItalic = False
'        .FontStrikeOut = False
         .PosX = 28
    End With
    With CustomColor
        .ForeColor = RGB(16, 0, 16)
        .DefTextColor = vbRed ' vbBlack
        .HilightColor = RGB(0, 255, 0)
        .NormalColor = RGB(186, 186, 204)
        .BackColor = RGB(58, 110, 165)
        .SelectedTextColor = RGB(0, 0, 255)
        .MenuTextColor = vbBlack
        .BorderColor = RGB(240, 72, 72)
    End With
    With CustomMenu
        .Icon.Add Array(101, 102), "Open"
        .Icon.Add Array(105, 106), "Mute"
        .Icon.Add Array(107, 108), "Play"
        .Icon.Add Array(113, 114), "Pause"
        .Icon.Add Array(109, 110), "FForward"
        .Icon.Add Array(111, 112), "FRewind"
        .Icon.Add Array(115, 116), "Tool"
        .Icon.Add Array(119, 120), "Hide Desktop"
        .Icon.Add Array(117, 118), "Show Desktop"
        .Icon.Add Array(121, 122), "Stop"
    End With
    SetIconTray Form1.imgIcon, "VIDEO DESKTOP"

    With Me
        oldProc = SetWindowLongA(.hwnd, GWL_WNDPROC, AddressOf WndProc)
        SetHotKey .hwnd, MOD_CONTROL, Asc("O")
        SetHotKey .hwnd, MOD_SHIFT, vbKeyDelete
        SetHotKey .hwnd, MOD_SHIFT, vbKeyInsert
        SetHotKey .hwnd, MOD_CONTROL, vbKeyLeft
        SetHotKey .hwnd, MOD_CONTROL, vbKeyRight
        SetHotKey .hwnd, MOD_CONTROL, Asc("P")
        SetHotKey .hwnd, MOD_CONTROL, Asc("S")
    End With 'Me
    On Error GoTo 0

End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           Y As Single)
    lngMsg = x / Screen.TwipsPerPixelX
    If Not blnFlag Then
        blnFlag = True
        Select Case lngMsg
        Case WM_RBUTTONUP
            SetForegroundWindow Me.hwnd
            Me.PopupMenu mnuOp
        End Select
        blnFlag = False
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    SetWindowLongA Me.hwnd, GWL_WNDPROC, oldProc
    stopODMenus Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WALLPAPERTHEME (RestoreWall)
    CLOSEALLPLAYER
    RemoveIconTray Form1.imgIcon, "VIDEO DESKTOP"
    ReleasHotKey Me.hwnd
End Sub
Public Sub HDesktop()
    mnuHD_Click
End Sub
Private Sub mnuExit_Click()
    CLOSEALLPLAYER
    Unload Me
End Sub
Private Sub mnuFF_Click()
    MM.FForward 5 '5 seconds
End Sub
Private Sub mnuFName_Click()
tmrLoop.Enabled = False
    If isPlaying Then
        CLOSEALLPLAYER
        isPlaying = False
    End If
    filmname = OpenDialog(Me.hwnd)
    If LenB(filmname) Then
       mnuPlay_Click
    End If
End Sub
Private Sub mnuFR_Click()
    MM.FRewind 5
End Sub
Private Sub mnuHD_Click()
    HideDesktop True
    mnuSD.Enabled = True
    mnuHD.Enabled = False
End Sub

Private Sub mnuloop_Click()
mnuloop.Checked = Not mnuloop.Checked
End Sub

Private Sub mnuM_Click()
    isMute = IIf(isMute, False, True)

    If isMute Then
        MM.SetAudioState Chan_All, vd_Off
    Else
        MM.SetAudioState Chan_All, vd_On
    End If
End Sub
Private Sub mnuPause_Click()
    MM.setCommand (PauseCD)
    isPaused = True
End Sub
Private Sub mnuPlay_Click()
On Error Resume Next
    If isPaused Then
        MM.setCommand (ResumeCD)
    Else
        With MM
            .Filename = filmname
            WALLPAPERTHEME (ClearWall)
            .PlayMEDIAFILE
        End With
    
    isPlaying = True
    isPaused = False
    InvalidateRect 0&, ByVal 0, 1&
    tmrLoop.Enabled = True
    tmrLoop.Interval = 500
    End If
    On Error GoTo 0
End Sub
Private Sub mnuSD_Click()
    HideDesktop False
    mnuHD.Enabled = True
    mnuSD.Enabled = False
    InvalidateRect 0&, ByVal 0, 1&
End Sub
Private Sub mnuSt_Click()
    mnuSt.Checked = Not mnuSt.Checked
    StarUP = IIf(StarUP, False, True)

    If StarUP Then
        SetStart False
    Else
        SetStart True
    End If
End Sub
Private Sub mnuStop_Click()
If isPlaying Then
    CLOSEALLPLAYER
    WALLPAPERTHEME (RestoreWall)
    isPlaying = False
    tmrLoop.Interval = 0
End If
End Sub
Public Sub OPENVIDEO()
    mnuFName_Click
End Sub
Public Sub Play()
    mnuPlay_Click
End Sub
Public Sub SDesktop()
    mnuSD_Click
End Sub
Public Sub StopPlay()
    mnuStop_Click
End Sub

Private Sub tmrLoop_Timer()
If MM.THE_ENDOFSONG(ByMS) Then
   If mnuloop.Checked Then
      mnuPlay_Click
   Else
      mnuStop_Click
   End If
End If
End Sub
