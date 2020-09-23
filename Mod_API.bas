Attribute VB_Name = "Mod_API"
Option Explicit

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize _
As Long) As Long

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal _
lpString As String, ByVal nCount As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) _
As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'{
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
    As Long
    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
        0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
              0, 0, FLAGS)
        SetTopMostWindow = False
    End If
End Function
