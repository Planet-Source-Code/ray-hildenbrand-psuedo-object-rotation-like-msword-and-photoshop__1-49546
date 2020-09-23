Attribute VB_Name = "modRotation"
Option Explicit

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Declare Function CreateCompatibleDC Lib "gdi32" _
        (ByVal hdc As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hdc As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long

Public Const CF_BITMAP      As Long = 2
Public Declare Function ReleaseDC Lib "user32" (ByVal HWND As Long, ByVal hdc As Long) As Long
Public Const SRCCOPY         As Long = &HCC0020
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

