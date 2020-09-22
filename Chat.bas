Attribute VB_Name = "chat"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_MINIMIZED = 6

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Const WM_DRAGFORM = &HA1

Public Declare Function ExitWindowsEx Lib "User32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' use in a MouseDown procedure

Public Function DragForm(TheForm As Form) As Boolean
    
    Call ReleaseCapture
    
    retval& = SendMessage(TheForm.hWnd, &HA1, 2, 0&)
    
End Function


