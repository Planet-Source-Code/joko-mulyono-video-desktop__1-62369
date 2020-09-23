Attribute VB_Name = "RegistryMod"
Option Explicit

    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
    End Type
    
    Public Enum Theme
        ClearWall
        StoreWall
        RestoreWall
    End Enum

    #If False Then
    Private ClearWall, StoreWall, RestoreWall
    #End If
    
    Private Const COLOR_BACKGROUND            As Integer = 1
    Private Const ERROR_SUCCESS = 0
    Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

    Private Const HKEY_CLASSES_ROOT = &H80000000
    Private Const HKEY_CURRENT_CONFIG = &H80000005
    Private Const HKEY_CURRENT_USER = &H80000001
    Private Const HKEY_DYN_DATA = &H80000006
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const HKEY_PERFORMANCE_DATA = &H80000004
    Private Const HKEY_USERS = &H80000003
    Private Const REG_SZ                      As Long = 1 ' Unicode nul terminated string
    Private Const REG_DWORD                   As Long = 4 ' 32-bit number
    Private Const REG_EXPAND_SZ               As Long = 2 'Unicode nul terminated string
    Private Const KEY_ALL_ACCESS              As Long = &H3F
    Private Const ERROR_NONE                  As Integer = 0
    Private Const SPI_SETDESKWALLPAPER        As Integer = 20
    Private Const SPIF_SENDWININICHANGE       As Long = &H2
    Private Const SPIF_UPDATEINIFILE          As Long = &H1
    Private ppt_retWall                       As String
    Private ppt_retStyle                      As String
    Private retOrig                           As String
'---------------------------------------------------------------
'-Registry API Declarations...
'---------------------------------------------------------------

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal Reserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                ByVal dwOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                ByRef phkResult As Long, _
                                                                                ByRef lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal lpReserved As Long, _
                                                                                  ByRef lpType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                            ByVal lpValueName As String, _
                                                                                            ByVal lpReserved As Long, _
                                                                                            lpType As Long, _
                                                                                            ByVal lpData As String, _
                                                                                            lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          ByVal lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByVal lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SetSysColors Lib "user32.dll" (ByVal nChanges As Long, _
                                                        lpSysColor As Long, _
                                                        lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                ByVal hWndNewParent As Long) As Long
Private Sub ClearDesktop()
    SetSysColors 1, COLOR_BACKGROUND, RGB(16, 0, 16)
    UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", ""
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, "", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub
Public Function getString(Str As String) As String
Dim a As Integer
    For a = 1 To Len(Str)
        If Mid$(Str, a, 1) = vbNullChar Then
            Exit For
        End If
    Next a
    getString = RTrim$(left(Str, a - 1))
End Function

Public Sub SetStart(ByVal StartUp As Boolean)

    If StartUp Then
    
        UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe C:\WINDOWS\VIDEODESKTOP.exe"
    Else
        UpdateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", ""
    End If
End Sub
Private Sub SetWallPaper(ByVal Display As Integer, _
                         ByVal sdir As String)
'TESTED:OK
Dim NewPaper As String
    NewPaper = sdir
    Select Case Display
    Case 0
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"

    Case 1
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"

    Case 2
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        UpdateKey HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2"

    End Select
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, NewPaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

End Sub
Public Sub StoreWallpaper()
'TESTED:OK
    On Error Resume Next
    ppt_retWall = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "Wallpaper")
    ppt_retStyle = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "WallpaperStyle")
    retOrig = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\", "OriginalWallpaper")
    If LenB(ppt_retWall) = 0 Then
        ppt_retWall = retOrig
    ElseIf ppt_retStyle = Not IsNumeric(ppt_retStyle) Then
        ppt_retStyle = 2 'as default [ stretch ]
    End If
    On Error GoTo 0
End Sub
Public Sub WALLPAPERTHEME(thmOpt As Theme)
'TESTED:OK
    Select Case thmOpt
    Case ClearWall
        StoreWallpaper
        ClearDesktop
    Case StoreWall
        StoreWallpaper
    Case RestoreWall
        If Not IsNumeric(ppt_retStyle) Then
            ppt_retStyle = 2 'as default [ stretch ]
        End If
        SetWallPaper ppt_retStyle, ppt_retWall
'SetSysColors 1, COLOR_BACKGROUND, Form1.Image1.BackColor 'RGB(16, 0, 16)
    End Select
End Sub


''Public Function CheckStart() As String
''CheckStart = getString(QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"))
''End Function

''Public Function GetWindowsColor() As OLE_COLOR
''GetWindowsColor = GetSysColor(COLOR_BACKGROUND)
''End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
      
    tmpVal = left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
    
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...


    'Create/Open Registry Key

    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
    

    ' Create/Modify Key Value

    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...
    
    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
  
    'Close Registry Key
    rc = RegCloseKey(hKey)                              ' Close Key
    UpdateKey = True                                    ' Return Success
    Exit Function                                       ' Exit
CreateKeyError:
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
End Function

