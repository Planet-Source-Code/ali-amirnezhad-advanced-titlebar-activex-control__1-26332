Attribute VB_Name = "modMouseOver"
Option Explicit

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

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Function IsMouseInto(ByVal lngHWnd As Long) As Boolean
  Dim btnMousePosition As POINTAPI
  Dim RetValue As Boolean
  Dim rctButton As RECT
  
  Call GetCursorPos(btnMousePosition)
  Call GetWindowRect(lngHWnd, rctButton)
  
  IsMouseInto = False
  If ((btnMousePosition.X >= rctButton.Left) And (btnMousePosition.X <= rctButton.Right)) And ((btnMousePosition.Y >= rctButton.Top) And (btnMousePosition.Y <= rctButton.Bottom)) Then
    IsMouseInto = True
  End If
End Function
