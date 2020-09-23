VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ==================================================================================
' cDIBSection.cls
' Copyright 1998-2003 Steve McMahon (steve@vbaccelerator.com)
' Visit vbAccelerator at www.vbaccelerator.com
'
' Creates and manages a True Color (24bit) GDI DibSection.
' This is DIB in which the bitmap bits are stored in windows
' memory so can be modified.
' See the RandomiseBits and Resample methods for how to do it.
'
' For example, fading in an out a 256x256 true colour DIB by
' directly modifying the bytes runs at 38fps on my machine
' (PII 266Mhz, 32Mb RAM, 8Mb ATI Xpert@Work AGP card)
'
' Note: for best performance, when compiling an executable check
' all the boxes on the Properties-Compile tab Advanced Optimisations
' button.  This really makes a difference! (e.g. the fading example
' ran at 22fps before I did this so > 50%!).
' ==================================================================================
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
   Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
   ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hDestDC As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
   ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
   ByVal hInst As Long, ByVal lpsz As String, _
   ByVal un1 As Long, _
   ByVal n1 As Long, ByVal n2 As Long, _
   ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
   ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Start of structure:
Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
   bfType As Integer '- type  ="BM" i.e &H4D42 - 2
   bfSize As Long ' - size in bytes of file - 6
   bfReserved1 As Integer ' - reserved, must be 0 - 8
   bfReserved2 As Integer ' - reserved, must be 0 - 10
   bfOffBits As Long ' offset from this structure to the bitmap bits - 14
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
   ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
   ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" ( _
   ByVal hFile As Long, lpBuffer As Any, _
   ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, _
   lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" ( _
   ByVal hFile As Long, _
   lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
   lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" ( _
   ByVal hFile As Long, _
   ByVal lDistanceToMove As Long, _
   lpDistanceToMoveHigh As Long, _
   ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
   ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
   ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
   ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
 ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
 ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
 ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" ( _
   ByVal aHDC As Long, ByVal hBitmap As Long, _
   ByVal nStartScan As Long, ByVal nNumScans As Long, _
   lpBits As Long, lpbi As BITMAPINFO, _
   ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" ( _
   ByVal hDC As Long, _
   lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, _
   lpInitBits As Any, lpInitInfo As BITMAPINFO, _
   ByVal wUsage As Long) As Long

' DrawDIB functions:
Private Declare Function DrawDibOpen Lib "msvfw32.dll" () As Long
Private Declare Function DrawDibClose Lib "msvfw32.dll" ( _
   ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "msvfw32.dll" ( _
   ByVal hDD As Long, _
   ByVal hDC As Long, _
   ByVal xDst As Long, ByVal yDst As Long, _
   ByVal dxDst As Long, ByVal dyDst As Long, _
   lpbi As Any, lpBits As Any, _
   ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dxSrc As Long, ByVal dySrc As Long, _
      ByVal wFlags As Long) As Long

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "Msimg32.dll" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long
Private Declare Function TransparentBlt Lib "Msimg32.dll" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal hHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal crTransparent As Long _
 ) As Long

Private m_hDIB As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_hDD As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO

Public Property Get UseDrawDib() As Boolean
   UseDrawDib = Not (m_hDD = 0)
End Property
Public Property Let UseDrawDib(ByVal bState As Boolean)
   If bState Then
      If m_hDD = 0 Then
         m_hDD = DrawDibOpen()
      End If
   Else
      If Not (m_hDD = 0) Then
         DrawDibClose m_hDD
      End If
   End If
End Property


Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDIB As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanline * .biHeight
    End With
    hDIB = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDIB <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = GetDC(GetDesktopWindow())
        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteObject lhDC
            End If
        End If
    End If
End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
Dim bDrawDib As Boolean
   bDrawDib = UseDrawDib()
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIB)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIB)
            UseDrawDib = bDrawDib
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanline() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanline = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lhDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub
Public Function SavePicture(ByVal sFilename As String) As Boolean
Dim lC As Long, i As Long

   ' Save to BMP:
   SavePicture = SaveToBitmap(m_lPtr, sFilename)

End Function
Private Function SaveToBitmap(ByVal lPtrBits As Long, ByVal sFilename As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lPtr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(m_tBI)
      .bfSize = .bfOffBits + m_tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFilename, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = Err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      Err.Raise 17, App.EXEName & ".cDIBSection", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lPtr = GlobalLock(hMem)
      CopyMemory ByVal lPtr, tBH.bfType, 2
      CopyMemory ByVal lPtr + 2, tBH.bfSize, 4
      CopyMemory ByVal lPtr + 6, 0&, 4
      CopyMemory ByVal lPtr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lPtr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(m_tBI)
         lR = WriteFile(hFile, m_tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = m_tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      SaveToBitmap = Not (bErr)
   End If

End Function
Private Function ApiError(ByVal E As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, E, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function
Private Function FileErrHandler( _
      ByVal lR As Long, _
      ByVal lSize As Long, ByVal lBytes As Long _
   ) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function

Public Sub CopyToClipboard()

  Dim dsk_Wnd     As Long
  Dim dsk_hDC     As Long

  Dim clp_hDC     As Long
  Dim clp_hObj    As Long
  Dim clp_hOldObj As Long

    If (Me.hDIB <> 0) Then

        If (OpenClipboard(0) <> 0) Then

            dsk_Wnd = GetDesktopWindow
            dsk_hDC = GetDC(dsk_Wnd)

            clp_hDC = CreateCompatibleDC(dsk_hDC)
            If (clp_hDC <> 0) Then

                clp_hObj = CreateCompatibleBitmap(dsk_hDC, Me.Width, Me.Height)
                If (clp_hObj <> 0) Then

                    clp_hOldObj = SelectObject(clp_hDC, clp_hObj)
                    Me.Paint clp_hDC
                    SelectObject clp_hDC, clp_hOldObj

                    EmptyClipboard
                    SetClipboardData CF_BITMAP, clp_hObj
                    CloseClipboard

                    DeleteDC clp_hDC
                End If
            End If

            ReleaseDC dsk_Wnd, dsk_hDC
        End If
    End If
End Sub
Public Sub Paint(ByVal hDstDC As Long, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, Optional ByVal ScaleFactor As Single = 1)
    
  Dim lret As Long
  
    If (Me.hDIB <> 0) Then
        With m_tBI.bmiHeader
            lret = StretchBlt(hDstDC, X, Y, .biWidth * ScaleFactor, -.biHeight * ScaleFactor, m_hDC, 0, 0, .biWidth, -.biHeight, SRCCOPY)
        End With
    End If
End Sub
Public Sub PaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy, _
        Optional ByVal crTransparent As Long = -1 _
    )
   If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
   If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
   If Not (m_hDD = 0) Then
      ' DrawDib method:
      DrawDibDraw m_hDD, lhDC, lDestLeft, lDestTop, _
      lDestWidth, lDestHeight, _
       m_tBI, _
       ByVal m_lPtr, _
       lSrcLeft, lSrcTop, _
       lDestWidth, lDestHeight, 0
   Else
      If Not (crTransparent = -1) Then
         TransparentBlt lhDC, lDestLeft, lDestTop, _
            lDestWidth, lDestHeight, _
            m_hDC, _
            lSrcLeft, lSrcTop, _
            lDestWidth, lDestHeight, _
            crTransparent
      Else
         BitBlt lhDC, lDestLeft, lDestTop, _
            lDestWidth, lDestHeight, _
            m_hDC, _
            lSrcLeft, lSrcTop, eRop
      End If
   End If
End Sub
Public Sub AlphaPaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lConstantAlpha As Byte = 255 _
    )
   If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
   If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
   
   Dim lBlend As Long
   Dim bf As BLENDFUNCTION
   bf.BlendOp = AC_SRC_OVER
   bf.BlendFlags = 0
   bf.SourceConstantAlpha = lConstantAlpha
   bf.AlphaFormat = AC_SRC_ALPHA
   CopyMemory lBlend, bf, 4
   
   Dim lR As Long
   lR = AlphaBlend( _
      lhDC, _
      lDestLeft, lDestTop, lDestWidth, lDestHeight, _
      m_hDC, _
      lSrcLeft, lSrcTop, lDestWidth, lDestHeight, _
      lBlend)
   If (lR = 0) Then
      Debug.Print ApiError(Err.LastDllError)
   End If

End Sub
Public Property Get hDC() As Long
    hDC = m_hDC
End Property
Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim lC As Long
Dim tSA As SAFEARRAY2D
Dim xEnd As Long
    
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanline()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:
    Randomize Timer
    
    xEnd = (Width - 1) * 3
    If (bGray) Then
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            For X = 0 To xEnd Step 3
                lC = Rnd * 255
                bDib(X, Y) = lC
                bDib(X + 1, Y) = lC
                bDib(X + 2, Y) = lC
            Next X
        Next Y
    Else
        For X = 0 To xEnd Step 3
            For Y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(X, Y) = 0
                bDib(X + 1, Y) = Rnd * 255
                bDib(X + 2, Y) = Rnd * 255
            Next Y
        Next X
    End If
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIB
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0: m_hDIB = 0: m_hBmpOld = 0: m_lPtr = 0
    If Not (m_hDD = 0) Then
      DrawDibClose m_hDD
      m_hDD = 0
   End If
End Sub

Public Function Resample( _
        ByVal lNewWidth As Long, _
        Optional ByVal lNewHeight As Long = -1 _
    ) As cDIBSection
   
   If (lNewHeight = -1) Then
      lNewHeight = (Height * lNewWidth) \ Width
   End If
   
   Dim cDib As cDIBSection
   Set cDib = New cDIBSection
   If cDib.Create(lNewWidth, lNewHeight) Then
      If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or _
         (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
         ' Change in size, do resample:
         ResampleDib cDib
      Else
         ' No size change so just return a copy:
         cDib.LoadPictureBlt m_hDC
      End If
      Set Resample = cDib
   End If
   
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
Dim bDibFrom() As Byte
Dim bDibTo() As Byte
Dim tSAFrom As SAFEARRAY2D
Dim tSATo As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanline()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanline()
        .pvData = cDibTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

Dim xScale As Single
Dim yScale As Single

Dim X As Long, Y As Long, xEnd As Long, xOut As Long

Dim fX As Single, fY As Single
Dim ifY As Long, ifX As Long
Dim DX As Single, DY As Single
Dim r As Long, R1 As Single, R2 As Single, r3 As Single, r4 As Single
Dim g As Long, G1 As Single, G2 As Single, g3 As Single, g4 As Single
Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
Dim ir1 As Long, ig1 As Long, ib1 As Long
Dim ir2 As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For Y = 0 To cDibTo.Height - 1
        
        fY = Y * yScale
        ifY = Int(fY)
        DY = fY - ifY
        
        For X = 0 To xEnd
            fX = X * xScale
            ifX = Int(fX)
            DX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY)
            G1 = bDibFrom(ifX + 1, ifY)
            R1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY)
            G2 = bDibFrom(ifX + 4, ifY)
            R2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1)
            g3 = bDibFrom(ifX + 1, ifY + 1)
            r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1):
             r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = R1 * (1 - DY) + r3 * DY
            ig1 = G1 * (1 - DY) + g3 * DY
            ib1 = b1 * (1 - DY) + b3 * DY
            ir2 = R2 * (1 - DY) + r4 * DY
            ig2 = G2 * (1 - DY) + g4 * DY
            ib2 = b2 * (1 - DY) + b4 * DY
            ' Interpolate in y:
            r = ir1 * (1 - DX) + ir2 * DX
            g = ig1 * (1 - DX) + ig2 * DX
            b = ib1 * (1 - DX) + ib2 * DX
            
            ' Set output:
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            If (b < 0) Then b = 0
            If (b > 255) Then
                b = 255
            End If
            xOut = X * 3
            bDibTo(xOut, Y) = b
            bDibTo(xOut + 1, Y) = g
            bDibTo(xOut + 2, Y) = r
            
        Next X
        
    Next Y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function

Private Sub Class_Terminate()
    ClearUp
End Sub


