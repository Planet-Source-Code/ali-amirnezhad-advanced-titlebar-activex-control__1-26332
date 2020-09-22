Attribute VB_Name = "modApi"
Option Explicit

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub DragParentForm(ByVal hndObject As Long)
  ReleaseCapture
  Call SendMessage(hndObject, &HA1, 2, 0&)
End Sub

