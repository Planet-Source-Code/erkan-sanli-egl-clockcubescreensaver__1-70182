Attribute VB_Name = "modMain"
Option Explicit

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2

Private Const SPI_SCREENSAVERRUNNING = 97&

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private OSVer As OSVERSIONINFO


Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_CHILD = &H40000000
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)
Private rctThumb As RECT

'Public Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hWnd&) As Boolean
'Private Declare Function PwdChangePassword Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname As String, ByVal hWnd As Long, ByVal uiReserved1 As Long, ByVal uiReserved2 As Long) As Long
Public Thumbnail As Boolean

Sub Main()
    
    On Error GoTo ErrorHandle
    
    Call GetVersionEx(OSVer)
    Select Case LCase(Left(Command, 2))
        Case "/s": Call StartScreenSaver    'Screensaver "S" TART
        Case "/c": frmSetup.Show            'Screensaver "C" ONFIGURE
        Case "/p": Call StartThumbnail      'Thumbnail   "P" REVIEW
        Case Else: frmPanel.Show
    End Select
    Exit Sub

ErrorHandle:
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
    Unload frmCanvas
End Sub

Public Function FileExists(FullFileName As String) As Boolean
    
    On Error GoTo FileExistsError
    If UCase(Left(FullFileName, 10)) = "NO PICTURE" Then GoTo FileExistsError
    Open FullFileName For Input As #1
    Close #1
    FileExists = True           'file exists
    Exit Function
    
FileExistsError:
    FileExists = False          'file does not exist
    Exit Function
    
End Function

Public Sub StartScreenSaver()
    
    Dim wFlags As Long
    Dim RetVal As Long
    
    If App.PrevInstance Then End
    
    Call RegRead
    If Params.First = False Then
        frmSetup.Show vbModal
        Call RegRead
    Else
'Hide cursor
        ShowCursor 0
'frmCanvas on top
        wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
        SetWindowPos frmCanvas.hWnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags
'Disable Control+Alt+Delete and Alt+Tab
        If OSVer.dwPlatformId <> VER_PLATFORM_WIN32_NT Then
            RetVal = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, 0&, 0&)
        End If
        Thumbnail = False
        frmCanvas.Show
    End If
    
End Sub

Public Sub EndScreenSaver()
    
    Dim RetVal As Long

'Show cursor
    ShowCursor 1
'Enable Control+Alt+Delete
    If OSVer.dwPlatformId <> VER_PLATFORM_WIN32_NT Then
        RetVal = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, 0&, 0&)
    End If
    
End Sub

Private Sub StartThumbnail()

    Dim lStyle As Long
    Dim hThumb As Long
    
    On Error GoTo ErrorHandle
    
    hThumb = CLng(Right(Command, Len(Command) - 2))
    Thumbnail = True
    GetClientRect hThumb, rctThumb
    lStyle = GetWindowLong(frmCanvas.hWnd, GWL_STYLE)
    lStyle = lStyle Or WS_CHILD
    SetWindowLong frmCanvas.hWnd, GWL_STYLE, lStyle
    SetParent frmCanvas.hWnd, hThumb
    SetWindowLong frmCanvas.hWnd, GWL_HWNDPARENT, hThumb
    SetWindowPos frmCanvas.hWnd, HWND_TOP, 0, 0, rctThumb.Right, rctThumb.Bottom, SWP_SHOWWINDOW
    Exit Sub

ErrorHandle:

    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
    Unload frmCanvas

End Sub

