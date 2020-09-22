Attribute VB_Name = "modGradient"
Option Explicit

Public Type rgbColor
  lngRed As Long
  lngBlue As Long
  lngGreen As Long
End Type
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function GetRGBColor(ByVal lngColor As Long) As rgbColor
  Dim dblBlue As Double
  Dim dblGreen As Double
  Dim dblRed As Double
  
  dblBlue = Fix((lngColor / 256) / 256)
  dblGreen = Fix((lngColor - ((dblBlue * 256) * 256)) / 256)
  dblRed = Fix(lngColor - ((dblBlue * 256) * 256) - (dblGreen * 256))
  GetRGBColor.lngRed = dblRed
  GetRGBColor.lngBlue = dblBlue
  GetRGBColor.lngGreen = dblGreen
End Function

