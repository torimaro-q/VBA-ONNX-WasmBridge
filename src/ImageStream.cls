VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, ByRef pbmi As BITMAPINFO, ByVal iUsage As Long, ByRef ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBM As LongPtr, ByVal hpal As LongPtr, Bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, ByRef clsidEncoder As GUID, ByVal encoderParams As Any) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpszCLSID As LongPtr, ByRef pCLSID As GUID) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, outputbuf As Any) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef Bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As LongPtr, ByRef width As Long) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As LongPtr, ByRef height As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal Bitmap As LongPtr, ByRef Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal Bitmap As LongPtr, lockedBitmapData As BitmapData) As Long
Private Declare PtrSafe Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal length As Long)
Private udtInput As GdiplusStartupInput
Private bmpdata As BitmapData
Private bsize As Long
Public Property Get width() As Long
    width = bmpdata.width
End Property
Public Property Get height() As Long
    height = bmpdata.height
End Property
Public Function GetBmpStreamFromFile(ByVal FileName As String) As Byte()
    Dim scan0 As LongPtr, lngToken As LongPtr, pSrcBmp As LongPtr
    Dim buffer() As Byte, IRect As Rect, w As Long, h As Long
    udtInput.GdiplusVersion = 1
    If GdiplusStartup(lngToken, udtInput, ByVal 0&) <> 0 Then End
    If GdipCreateBitmapFromFile(ByVal StrPtr(FileName), pSrcBmp) <> 0 Then GoTo err
    GdipGetImageWidth pSrcBmp, w
    GdipGetImageHeight pSrcBmp, h
    IRect.Right = w
    IRect.Bottom = h
    bsize = w * h * 4 - 1
    ReDim buffer(bsize) As Byte
    GdipBitmapLockBits pSrcBmp, IRect, &H3, PixelFormat32bppARGB, bmpdata
    scan0 = bmpdata.scan0
    MoveMemory VarPtr(buffer(0)), scan0, bsize
    GdipBitmapUnlockBits pSrcBmp, bmpdata
    GetBmpStreamFromFile = buffer
    GdipDisposeImage pSrcBmp
    GdiplusShutdown lngToken
    Exit Function
err:
    GdiplusShutdown lngToken
End Function

Private Function Stream2Hbitmap(ByRef source() As Byte, pw As Long, ph As Long) As LongPtr
    Dim hDIB As LongPtr, hbitmap As LongPtr, pBits As LongPtr, cbSize As Long
    Dim bmi As BITMAPINFO
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = pw
        .biHeight = -ph
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        .biSizeImage = pw * ph * 4
    End With
    hbitmap = CreateDIBSection(0, bmi, 0, pBits, 0, 0)
    If hbitmap = 0 Then Exit Function
    cbSize = pw * ph * 4
    MoveMemory pBits, VarPtr(source(0)), cbSize
    Stream2Hbitmap = hbitmap
End Function
Private Function ConvCLSID(ByVal sGuid As String) As GUID
    Call CLSIDFromString(StrPtr(sGuid), ConvCLSID)
End Function
Public Sub SaveAsFile(ByRef source() As Byte, ByVal pw As Long, ByVal ph As Long, ByVal fname As String)
    Dim hbitmap As LongPtr
    hbitmap = Stream2Hbitmap(source, pw, ph)
    Dim sFormat As String, objBmp, depth As Long, udtInput  As GdiplusStartupInput, strEncoder As String, nQuarity As Long, uEncoderParams As EncoderParameters
    sFormat = StrConv(StrConv(Right(fname, 3), vbUpperCase), vbNarrow)
    Select Case sFormat
        Case "JPG": strEncoder = ENCODER_JPG
        Case "PNG": strEncoder = ENCODER_PNG
        Case Else: strEncoder = ENCODER_BMP
    End Select
    If sFormat = "JPG" Then
        nQuarity = 100
        uEncoderParams.Count = 1
        With uEncoderParams.Parameter(0)
            .IID = ConvCLSID(ENCODER_JPG)
            .Type = 4
            .NumberOfValues = 1
            .Value = CLng(VarPtr(nQuarity))
        End With
    End If
    Call GdipCreateBitmapFromHBITMAP(hbitmap, 0&, objBmp)
    If sFormat = "JPG" Then
        Call GdipSaveImageToFile(objBmp, StrPtr(fname), ConvCLSID(strEncoder), VarPtr(uEncoderParams))
        DoEvents
    Else
        Call GdipSaveImageToFile(objBmp, StrPtr(fname), ConvCLSID(strEncoder), ByVal 0&)
        DoEvents
    End If
    Call GdipDisposeImage(objBmp)
End Sub

