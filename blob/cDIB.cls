VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'*  VB file:   cDIB.cls... by Ray Mercer
'*  created:   12/1999 by Ray Mercer
'*  uploaded:  2/2000
'*  modified:  2/25/2000 by Ray Mercer
'*             Patrick Pasteels pointed out a bug in my code
'*             -fixed: ReDim m_memBitmapInfo(0 To 39) now correctly equals 40 bytes
'*  Copyright (C) 1999 - 2000 Ray Mercer.  All rights reserved.
'*  Latest version can be downloaded from http://www.shrinkwrapvb.com
'*-----------------------------------------------------------------
'*
'*  Slight modification by Jefri bin Mustapa
'*  Date : 03rd sept 2007
'*  Addition : GetObject , GetBitmapBits , SetBitmapBits , GetDIBits ,
'*             StretchDIBits from Lib "gdi32"
'*  Function Added :GetImageData(),SetImageData()
'*  Function Modified = CreateFromFile()
'****************************************************************
'*  some more Addition by Jefri bin Mustapa
'*  Date : 06th nov 2007
'*  Function Added : CropImData()
'*  Utilization : To get the image data of certain region selected in _
'*                a picture
'****************************************************************
Option Explicit

Private Const BMP_MAGIC_COOKIE As Integer = 19778 'this is equivalent to ascii string "BM"
'//BITMAP DEFINES (from mmsystem.h)
Private Type BITMAPFILEHEADER '14 bytes
        bfType As Integer '"magic cookie" - must be "BM"
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
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

Private Type RGBQUAD
    Red As Byte
    Green As Byte
    Blue As Byte
    Reserved As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'/* constants for the biCompression field */
Private Const BI_RGB  As Long = 0&
'#define BI_RLE8       1L
'#define BI_RLE4       2L
'#define BI_BITFIELDS  3L
'for use with AVIFIleInfo

'Private Type AVI_FILE_INFO  '108 bytes?
'    dwMaxBytesPerSecond As Long
'    dwFlags As Long
'    dwCaps As Long
'    dwStreams As Long
'    dwSuggestedBufferSize As Long
'    dwWidth As Long
'    dwHeight As Long
'    dwScale As Long
'    dwRate As Long
'    dwLength As Long
'    dwEditCount As Long
'    szFileType As String * 64
'End Type

'Private Declare Function CreateDIBSection_256 Lib "GDI32.DLL" Alias "CreateDIBSection" (ByVal hdc As Long, _
'                                                                                ByVal pbmi As BITMAPINFO_256, _
'                                                                                ByVal iUsage As Long, _
'                                                                                ByRef ppvBits As Long, _
'                                                                                ByVal hSection As Long, _
'                                                                                ByVal dwOffset As Long) As Long 'hBitmap
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long 'handle
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long 'Pointer to mem
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long 'BOOL
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)

Private Const HEAP_ZERO_MEMORY As Long = &H8

Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER
Private tmpgray() As Byte
'------------------------------------------
' Addition (Declaration)
'------------------------------------------
Private Type GRAYcolor
    Red As Byte
End Type

Private Type BITMAPINFO
    bmHeader As BITMAPINFOHEADER
    bmColors(0 To 255) As RGBQUAD
End Type


Private Type GRAYINFO
    bmHeader As BITMAPINFOHEADER
    bmColors(0 To 255) As GRAYcolor
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
'The magical API DIB function calls (they're long!)
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


'-------------------------------------------
'              Function Added
'-------------------------------------------
Public Function GetImageData(ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    Dim arraywidth As Long
    
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.biSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.biPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.biBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.biCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from SrcPictureBox and put it into our 'bm' variable
    GetObject SrcPictureBox.Image, bmLen, bm
    
    arraywidth = (bm.bmWidth * 3) - 1
    arraywidth = arraywidth + (bm.bmWidth Mod 4)
    'Build a correctly sized array
    'ReDim imagedata(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    ReDim ImageData(0 To arraywidth, 0 To bm.bmHeight - 1)
    'Finish building the 'bmi' variable we want to pass to the GetDIBits call (the same one we used above)
    bmi.bmHeader.biWidth = bm.bmWidth
    bmi.bmHeader.biHeight = bm.bmHeight
    'Now that we've completely filled up the 'bmi' variable, we use GetDIBits to take the data from
    'SrcPictureBox and put it into the ImageData() array using the settings we specified in 'bmi'
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0
End Function

'Routine to set an image's pixel information from an array dimensioned (rgb, x, y)
Public Function SetImageData(ByRef DstPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    Dim tempwidth As Long, tempheight As Long
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.biSize = 40  'Size, in bytes, of the header (always 40)
    bmi.bmHeader.biPlanes = 1  'Number of planes (always one for this instance)
    bmi.bmHeader.biBitCount = 24  'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.biCompression = 0  'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    
    Dim bmLen As Long
    bmLen = Len(bm)
   
    'DstPictureBox.Width = (UBound(ImageData, 1) \ 3) + 5
    'DstPictureBox.Height = UBound(ImageData, 2) + 5
    
    'Get the picture box information from DstPictureBox and put it into our 'bm' variable
    GetObject DstPictureBox.Image, bmLen, bm

    'Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call
    '(continuing to use the 'bmi' we used above)
    
    
    bmi.bmHeader.biWidth = bm.bmWidth
    bmi.bmHeader.biHeight = bm.bmHeight
    'Now that we've built the temporary header, we use StretchDIBits to take the data from the
    'ImageData() array and put it into SrcPictureBox using the settings specified in 'bmi' (the
    'StretchDIBits call should be on one continuous line)
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    'Note: Always set AutoRedraw to true when using DIB sections; when AutoRedraw is false
    'you will get unpredictable results.
    
    DstPictureBox.Picture = DstPictureBox.Image
    DstPictureBox.Refresh
   
End Function


Public Function CreateFromFile(ByVal filename As String, ByRef ImageData() As Byte) As Boolean
  Dim hFile As Long
    Dim tempheight As Long, tempwidth As Long
    Dim x As Long, y As Long
    Dim i As Long
        
    'If Not ExistFile(filename) Then
    '    MsgBox "File does not exist:" & vbCrLf & filename, vbCritical, App.title
    '    Exit Function
    'End If
   
    hFile = FreeFile()
    
     Open filename For Binary Access Read As #hFile
    
    'OK, file is opened - now for the real algorithm...
    Get #hFile, , m_bfh 'get the BITMAPFILEHEADER this identifies the bitmap

    If m_bfh.bfType <> BMP_MAGIC_COOKIE Then 'this is not a BMP file
        MsgBox "File is not a supported bitmap format:" & vbCrLf & filename, vbInformation, App.Title
        Close #hFile
        Exit Function
    Else
        'now get the info header
        Get #hFile, Len(m_bfh) + 1, m_bih 'start at the 15th byte
        
        'now get the bitmap bits
        ReDim m_memBits(0 To m_bih.biSizeImage - 1)
        Get #hFile, m_bfh.bfOffBits + 1, m_memBits
        
        'and BitmapInfo variable-length UDT
        ReDim m_memBitmapInfo(0 To m_bfh.bfOffBits - 14) 'don't need first 14 bytes (fileinfo)
        Get #hFile, Len(m_bfh) + 1, m_memBitmapInfo
        
        Close #hFile   'Close file
    End If
    
   
    
    tempheight = m_bih.biHeight - 1
    tempwidth = m_bih.biWidth - 1
    ReDim ImageData(0 To 2, 0 To tempwidth, 0 To tempheight)
    i = m_bih.biSizeImage - 1
    'i = 0
    For y = 0 To tempheight
        For x = 0 To tempwidth
        ImageData(2, tempwidth - x, tempheight - y) = m_memBits(i)
        i = i - 1
        ImageData(1, tempwidth - x, tempheight - y) = m_memBits(i)
        i = i - 1
        ImageData(0, tempwidth - x, tempheight - y) = m_memBits(i)
        i = i - 1
        
    Next
    Next
'    Debug.Print "BitCount: " & vbTab & vbTab & m_bih.biBitCount
'    Debug.Print "ClrImportant: " & vbTab & bih.biClrImportant
'    Debug.Print "ClrUsed: " & vbTab & vbTab & bih.biClrUsed
'    Debug.Print "Compression: " & vbTab & "&H" & Hex$(bih.biCompression)
'    Debug.Print "Height: " & vbTab & vbTab & bih.biHeight
'    Debug.Print "Planes: " & vbTab & vbTab & bih.biPlanes 'always 1
'    Debug.Print "Size: " & vbTab & vbTab & vbTab & m_bih.biSize
'    Debug.Print "SizeImage: " & vbTab & vbTab & m_bih.biSizeImage
'    Debug.Print "Width: " & vbTab & vbTab & vbTab & bih.biWidth
'    Debug.Print "XPelsPerMeter: " & vbTab & bih.biXPelsPerMeter 'usually 0
'    Debug.Print "YPelsPerMeter: " & vbTab & bih.biYPelsPerMeter 'usually 0
'<====ERROR TRAP ON
    On Error Resume Next
    
    'If Err Then
    '    If Err.Number = 70 Then
    '        MsgBox "File is locked - cannot access:" & vbCrLf & filename, vbCritical, App.title
    '    Else
    '        MsgBox Err.Description, vbInformation, App.title
    '    End If
         Exit Function 'assume file was not opened
    'End If
    'On Error GoTo 0
    '====>ERROR TRAP OFF
    
End Function

Public Function CreateFromPackedDIBPointer(ByRef pDIB As Long) As Boolean
On Error Resume Next
Debug.Assert pDIB <> 0
'Creates a full-color (no palette) DIB from a pointer to a full-color memory DIB

'get the BitmapInfoHeader
Call CopyMemory(ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih))
If m_bih.biBitCount < 16 Then
    Debug.Print "Error! DIB was less than 16 colors."
    Exit Function 'only supports high-color or full-color dibs
End If

'now get the bitmap bits
If m_bih.biSizeImage < 1 Then Exit Function 'return False
ReDim m_memBits(0 To m_bih.biSizeImage - 1)
Call CopyMemory(m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage)

'and BitmapInfo variable-length UDT
ReDim m_memBitmapInfo(0 To 39) 'don't need first 14 bytes (fileinfo)
Call CopyMemory(m_memBitmapInfo(0), m_bih, Len(m_bih))

'create a file header
With m_bfh
    .bfType = BMP_MAGIC_COOKIE
    .bfSize = 55 + m_bih.biSizeImage 'size of file as written to disk
    .bfReserved1 = 0&
    .bfReserved2 = 0&
    .bfOffBits = 54 'BitmapInfoHeader + BitmapFileHeader
End With

'and return True
CreateFromPackedDIBPointer = True

End Function

Public Function WriteToFile(ByVal filename As String) As Boolean
Dim hFile As Integer
On Error Resume Next
gray
DoEvents
hFile = FreeFile()
Open filename For Binary As hFile
    Put hFile, 1, m_bfh
    Put hFile, Len(m_bfh) + 1, m_memBitmapInfo
    Put hFile, , m_memBits
Close hFile
WriteToFile = True
End Function
Private Function gray()
    Dim i As Long, temp As Long
    For i = 0 To m_bih.biSizeImage - 1 Step 3
        temp = m_memBits(i)
        temp = temp + m_memBits(i + 1)
        temp = temp + m_memBits(i + 2)
        temp = temp \ 3
        m_memBits(i) = temp: m_memBits(i + 1) = temp: m_memBits(i + 2) = temp
    Next
End Function
Public Function allocGray(ByRef tmparry() As Byte, ByVal frame As Long)
' before using this u must first redim the tmparry
' give the current frame number, and total frames

Dim i As Long, temp As Long, j As Long
Dim tmpheight As Long, tmpwidth As Long
Dim y As Long, x As Long

tmpwidth = m_bih.biWidth - 1
tmpheight = m_bih.biHeight - 1

    For i = 0 To m_bih.biSizeImage - 1 Step 3
        temp = m_memBits(i)
        temp = temp + m_memBits(i + 1)
        temp = temp + m_memBits(i + 2)
        temp = temp \ 3
        m_memBits(i) = temp
    Next


  

i = m_bih.biSizeImage - 1
'i = 0
For y = 0 To tmpheight
    For x = 0 To tmpwidth
    
    'tmparry(frame, x, tmpheight - y) = m_memBits(y * (tmpwidth + 1) + (x * 3))
    tmparry(frame, tmpwidth - x, tmpheight - y) = m_memBits(i)
    i = i - 3
    Next
Next

    
End Function
Private Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (err = 0)
End Function

Public Property Get BitCount() As Long
    BitCount = m_bih.biBitCount

End Property

Public Property Get Height() As Long
    Height = m_bih.biHeight
End Property

Public Property Get Width() As Long
    Width = m_bih.biWidth
End Property

Public Property Get Compression() As Long
    Compression = m_bih.biCompression
End Property

Public Property Get SizeInfoHeader() As Long
    SizeInfoHeader = m_bih.biSize
End Property

Public Property Get SizeImage() As Long
    SizeImage = m_bih.biSizeImage
End Property

Public Property Get Planes() As Long
    Planes = m_bih.biPlanes
End Property

Public Property Get ClrImportant() As Long
    ClrImportant = m_bih.biClrImportant
End Property

Public Property Get ClrUsed() As Long
    ClrUsed = m_bih.biClrUsed
End Property

Public Property Get XPPM() As Long
    XPPM = m_bih.biXPelsPerMeter
End Property

Public Property Get YPPM() As Long
    YPPM = m_bih.biYPelsPerMeter
End Property

Public Property Get FileType() As Long
    FileType = m_bfh.bfType
End Property

Public Property Get SizeFileHeader() As Long
    SizeFileHeader = m_bfh.bfSize
End Property

Public Property Get BitOffset() As Long
    BitOffset = m_bfh.bfOffBits
End Property

Public Property Get PointerToBits() As Long
    PointerToBits = VarPtr(m_memBits(0))
End Property

Public Property Get PointerToBitmapInfo() As Long
    PointerToBitmapInfo = VarPtr(m_memBitmapInfo(0))
End Property

Public Property Get SizeBitmapInfo() As Long
    SizeBitmapInfo = UBound(m_memBitmapInfo()) + 1
End Property

Public Function GetImageDataStream(ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte)
'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    Dim arrayLen As Long
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.biSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.biPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.biBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.biCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from SrcPictureBox and put it into our 'bm' variable
    GetObject SrcPictureBox.Image, bmLen, bm
   
    'Build a correctly sized array
    'align the width so it wud be in times four..:)
    '------------------------------------------------
    arrayLen = bm.bmWidth * bm.bmHeight * 3
   
    ReDim ImageData(0 To arrayLen)
    '------------------------------------------------
    'Finish building the 'bmi' variable we want to pass to the GetDIBits call (the same one we used above)
    bmi.bmHeader.biWidth = bm.bmWidth
    bmi.bmHeader.biHeight = bm.bmHeight
    'Now that we've completely filled up the 'bmi' variable, we use GetDIBits to take the data from
    'SrcPictureBox and put it into the ImageData() array using the settings we specified in 'bmi'
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0), bmi, 0
    
End Function
Public Function CropImData(ByRef DstPictureBox As PictureBox, ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte, _
                           ByVal xmin As Long, ByVal xmax As Long, ByVal ymin As Long, ByVal ymax As Long, stretch As Boolean)
If stretch = False Then
   DstPictureBox.Width = xmax - xmin + 4
   DstPictureBox.Height = ymax - ymin + 4
End If
   
   StretchBlt DstPictureBox.hdc, 0, 0, DstPictureBox.ScaleWidth, DstPictureBox.ScaleHeight, SrcPictureBox.hdc, xmin, ymin, xmax - xmin, ymax - ymin, vbSrcCopy
   GetImageData DstPictureBox, ImageData
   SetImageData DstPictureBox, ImageData

End Function

