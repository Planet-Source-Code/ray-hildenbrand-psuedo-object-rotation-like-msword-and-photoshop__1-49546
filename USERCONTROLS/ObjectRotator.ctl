VERSION 5.00
Begin VB.UserControl ImageRotation 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ScaleHeight     =   3720
   ScaleWidth      =   5085
   ToolboxBitmap   =   "ObjectRotator.ctx":0000
   Begin VB.PictureBox picBuff2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   705
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   2850
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgRotate 
      Height          =   480
      Left            =   1155
      Picture         =   "ObjectRotator.ctx":0312
      Top             =   2745
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ImageRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////
' Original Concept By Ray Hildenbrand
''Copyright Ray Hildenbrand

' This code takes in to consideration three main concepts provided by other authors and
' extends their concepts together, namely
'   Steve McMahon Vbaccelerator.com - His code for DIB Processing and Region Mapping is superb fast and without his site, many of us would still be stuck on our first application! Thanks Steve
'   Zubuyer kaolin - His PsCode example on rotation (dial i believe) was used for most of the rotation logic, but was modified to handle the rotation point differently (top right as opposed to the center)
'   Florian Egal - His excellent library found on PSCode is the best! If you do not have this file (FoxCBmp3.dl) you will need to do a search on planetsourcecode for florian egal or advanced graphics 3.3 on copy the file to your sys32 directory.
''  If any one else sees any code that I did not mark as being theirs, please let me know and i will be happy to add their name.


''''Note, this code is not optimized....and is also a very small part of a much bigger project
'''' so there is a lot of code in here that is not really needed for the sake of this example, but i am not ripping it out to due time restrictions
'''' regardless, I though this was kinda neat so i thought i would share it with th vb community.
'''' Obviously, there needs some work on the rotation logic, as it is a little quirky but , you know.....

Option Explicit
Dim i
Dim j
Dim Degree
Dim bdeg
Dim mDIBRegion As cDIBSectionRegion
Dim LastDeg As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private mPic As StdPicture
Dim bMouseDown As Boolean

'for use in another app...blah blah, blah
Public LayerObject As Object
Event RotateCancelled(whichLayer As Object)
Event RotateRequestsFinalization(whichLayer As Object)
'///

Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Private Function DesignMode() As Boolean

On Local Error GoTo DesignMode_Error
    
    If UserControl.Ambient.UserMode Then
        DesignMode = False
    Else
        DesignMode = True
    End If
    Exit Function
    
DesignMode_Error:
    DesignMode = True
    Exit Function
End Function
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function
Private Sub pDrawRectangle(ByRef ObjectToDrawOn As Object, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
''''code from gonchuki
'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRect As RECT
Dim hBrush As Long
Dim ret As Long

bRect.Left = X
bRect.Top = Y
bRect.Right = X + Width
bRect.Bottom = Y + Height

hBrush = CreateHatchBrush(2, Color)

If OnlyBorder = False Then
    ret = FillRect(ObjectToDrawOn.hDC, bRect, hBrush)
Else
    ret = FrameRect(ObjectToDrawOn.hDC, bRect, hBrush)
End If
ret = DeleteObject(hBrush)

End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn And Not LayerObject Is Nothing Then
        RaiseEvent RotateRequestsFinalization(LayerObject)
    End If
    
    If KeyCode = vbKeyEscape And Not LayerObject Is Nothing Then
        RaiseEvent RotateCancelled(LayerObject)
    End If
        
    If KeyCode = vbKeyDelete Then
        UserControl.Extender.Visible = False
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim DX, DY, adeg, xdeg
Const PI = 3.14159265358979

If X > picBuffer.Width - 120 Then
    UserControl.MouseIcon = imgRotate.Picture
    UserControl.MousePointer = 99
Else
    UserControl.MousePointer = 0
End If
If Button Then

'The following codes do the main job
'It determines the degree of angle from the mouse coordinate
'modified from original authors source - Ray Hildenbrand 10/25/2003
'================================================================
        
        picBuff2.AutoRedraw = False
        Set picBuff2.Picture = LoadPicture()
        Dim tmpSquare As Long
        If picBuffer.Width > picBuffer.Height Then
            tmpSquare = picBuffer.Width + 750
        Else
            tmpSquare = picBuffer.Height + 750
        End If
        picBuff2.Move picBuffer.Left, picBuffer.Top, tmpSquare, tmpSquare
        picBuff2.Cls
       
        DX = (X - UserControl.Width / 2)
        DY = (Y - UserControl.Height / 2)
        If DX = 0 Then DX = 0.00001
        If DY = 0 Then DY = 0.00001
        
        xdeg = Atn(DY / DX)
        
        If DX > 0 And DY < 0 Then adeg = 6.283 - (xdeg * -1) '+,-
        If DX < 0 And DY < 0 Then adeg = 3.142 + (xdeg * 1) '-,-
        If DX < 0 And DY > 0 Then adeg = 3.142 - (xdeg * -1) '-,+
        If DX > 0 And DY > 0 Then adeg = (xdeg * 1) '+,+
        bdeg = (360 - (adeg * 180 / PI))
        
        'modified from original source for variance from original authors concept
        bdeg = Abs(bdeg - 360)
        
        bdeg = bdeg + 45
        
        If Fix(bdeg) = LastDeg Then
            Exit Sub
         Else
            DoEvents
         End If

        UserControl.Refresh
        
        DrawImageAtAngle Fix(bdeg)
        LastDeg = Fix(bdeg)
        
        
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Private Sub UserControl_Terminate()
    If Not mDIBRegion Is Nothing Then
        mDIBRegion.Applied(UserControl.hwnd) = False
        Set mDIBRegion = Nothing
    End If
    
End Sub
Public Property Get OriginalPicture() As StdPicture
    Set OriginalPicture = mPic
End Property
Public Property Set OriginalPicture(newPic As StdPicture)
    Set mPic = newPic
    Set UserControl.Picture = newPic
    Set picBuffer.Picture = newPic
    DrawImageAtAngle 0
    UserControl.Refresh
    
    If Not mDIBRegion Is Nothing Then
        mDIBRegion.Applied(UserControl.hwnd) = True
    End If
    UserControl.Refresh
End Property



Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    UserControl.Extender.Visible = vNewValue
End Property


Public Sub DrawImageAtAngle(Angle As Long, Optional DrawGrips As Boolean = True)
        If Not DesignMode Then
            UserControl.Extender.Visible = False
        End If
        picBuff2.AutoRedraw = False
        Set picBuff2.Picture = LoadPicture()
        Set picBuffer.Picture = mPic
        Dim tmpSquare As Long
        
        If picBuffer.Width > picBuffer.Height Then
            tmpSquare = picBuffer.Width + 750
        Else
            tmpSquare = picBuffer.Height + 750
        End If
        
        
        picBuff2.Move picBuffer.Left, picBuffer.Top, tmpSquare, tmpSquare
        picBuff2.Cls
        picBuffer.DrawWidth = 1
        picBuffer.DrawStyle = 4
        
        If DrawGrips Then
            pDrawRectangle picBuffer, 1, 1, picBuffer.ScaleWidth - 2, picBuffer.ScaleHeight - 2, vbBlack, True
            
            picBuffer.DrawStyle = 0
            picBuffer.DrawWidth = 1
            picBuffer.FillStyle = vbFSSolid
            picBuffer.FillColor = vbButtonFace
            picBuffer.Circle (3, 3), 3, vbBlack
            
            'topright
            picBuffer.FillColor = vbGreen
            'picBuffer.FillStyle = 4
            picBuffer.Circle (picBuffer.ScaleWidth - 5, 5), 4, &H800000
            
            picBuffer.FillColor = vbRed '&H44C2E8
            'picBuffer.FillStyle = 4
            picBuffer.Circle (picBuffer.ScaleWidth - 5, 5), 2, vbGreen
            
            picBuffer.FillColor = vbButtonFace
            'picBuffer.FillStyle = 4
            picBuffer.Circle (picBuffer.ScaleWidth - 4, picBuffer.ScaleHeight - 4), 3, &H8000000C
            picBuffer.FillColor = vbButtonFace
            picBuffer.Circle (3, picBuffer.ScaleHeight - 4), 3, &H8000000C
            Set picBuffer.Picture = picBuffer.Image
            
            
        End If
        UserControl.BackColor = vbWhite
        picBuffer.BackColor = vbWhite
        picBuff2.BackColor = vbWhite
        picBuff2.AutoRedraw = True
        picBuffer.FillStyle = vbFSSolid
        picBuff2.Refresh
        
        FoxRotate picBuff2.hDC, picBuff2.ScaleWidth / 2, picBuff2.ScaleHeight / 2, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBuffer.hDC, 0, 0, Fix(Angle), TranslateColor(vbWhite), FOX_ANTI_ALIAS + FOX_USE_MASK '''use you own transparent color eh
        picBuff2.Refresh
        Set picBuff2.Picture = picBuff2.Image
        picBuff2.AutoRedraw = False
        
        If mDIBRegion Is Nothing Then Set mDIBRegion = New cDIBSectionRegion
        Dim mDib As New cDIBSection
        
        Set UserControl.Picture = picBuff2.Picture
        mDib.CreateFromPicture picBuff2.Picture
        mDIBRegion.Create mDib, TranslateColor(vbWhite) ''use your own colors for transparency
        mDIBRegion.Applied(UserControl.hwnd) = True
        If Not UserControl.Width = mDib.Width * Screen.TwipsPerPixelX Then UserControl.Width = mDib.Width * Screen.TwipsPerPixelX
        If Not UserControl.Height = mDib.Height * Screen.TwipsPerPixelY Then UserControl.Height = mDib.Height * Screen.TwipsPerPixelY
        UserControl.Refresh
        Set picBuff2.Picture = Nothing
        UserControl.Refresh
        If Not DesignMode Then
            UserControl.Extender.Visible = True
        End If
        LastDeg = Angle
End Sub

Public Property Get RenderedPicture() As StdPicture
    DrawImageAtAngle CDbl(LastDeg), False
    Set RenderedPicture = UserControl.Picture
    DrawImageAtAngle CDbl(LastDeg), True
End Property
