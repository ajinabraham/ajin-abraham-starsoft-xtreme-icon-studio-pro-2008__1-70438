Attribute VB_Name = "Module1"
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'function for converting app coordinates to screencordiantes
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'function to get the title bar caption of a window
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'to get the window rect
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long



Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 







'############For transparent window
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_LAYERED = &H80000

Public Const LWA_ALPHA = &H2&


Public Sub SetTrans(ByVal alpha As Long, window As Long)
On Error GoTo warn
alpha = 255 - alpha
Dim lOldStyle As Long
    Dim LhWnd As Long
    
    LhWnd = window
    
        lOldStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
        SetWindowLong LhWnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes LhWnd, 0, alpha, LWA_ALPHA
Exit Sub

warn:
MsgBox "Select a suitable window first!", vbCritical, "Select a window"
End Sub
