Attribute VB_Name = "Util"
Public Type BITMAPINFOHEADER
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
Public Type RGBA
    R As Byte
    G As Byte
    B As Byte
    A As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0) As RGBA
End Type
Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type BitmapData
    width As Long
    height As Long
    Stride As Long
    PixelFormat As Long
    scan0 As LongPtr
    Reserved As Long
End Type
Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Public Enum PixelFormat
    PixelFormat32bppARGB = &H26200A
End Enum
Public Type PICTDESC
    cbSizeofStruct As Long
    picType As Long
    hbitmap As LongPtr
    hPalette As LongPtr
    Reserved As LongPtr
End Type
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Type EncoderParameter
    IID           As GUID
    NumberOfValues As Long
    Type           As Long
    Value          As Long
End Type
Public Type EncoderParameters
    Count         As Long
    Parameter(15) As EncoderParameter
End Type
Public Type L
    v As Long
End Type
Public Type Dtyp
    v As Double
End Type
Public Type Styp
    v As Single
End Type
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Const ENCODER_BMP    As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Public Const ENCODER_JPG    As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Public Const ENCODER_PNG    As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Public Const PICTYPE_BITMAP = 1
Public Const CLMASK As Long = &H80000000
Public Const LOGPIXELSX As Long = 88
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const DEG2RAD As Double = 0.017453293
Public Const Inv255 As Double = 1 / 255
Public Function ModelNames() As Variant
    ModelNames = Array("VbOnnxResNet", "VbOnnxYOLOX", "VbOnnxFNST", "VbOnnxMiDaS", "VbOnnxAnimeGAN", "VbOnnxOCR", "VbOnnxUF_FER")
End Function
Public Function json2coll(ByVal json As String) As Collection
    Dim buf As Variant: buf = Split(Replace(Replace(json, "]", ""), "[", ""), "{")
    Dim coll As Collection: Set coll = New Collection
    Dim B As Variant
    For Each B In buf
        If Len(B) > 1 Then
            Dim elms, e, k, v, spl
            Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
            With dic
                elms = Split(Replace(B, "}", ""), ",""")
                For Each e In elms
                    If Len(e) > 1 Then
                        spl = Split(e, ":")
                        k = Replace(CStr(spl(0)), """", "")
                        v = Replace(CStr(spl(1)), """", "")
                        If v Like "*,*?" Then
                            .Add k, Split(v, ",")
                        Else
                            If IsNumeric(v) Then
                                .Add k, CDbl(v)
                            Else
                                .Add k, v
                            End If
                        End If
                    End If
                Next e
            End With
            coll.Add dic
        End If
    Next B
    Set json2coll = coll
End Function
Public Function GetFileName(ByVal filepath As String, Optional dlm As String = "\") As String
    GetFileName = Split(filepath, dlm)(UBound(Split(filepath, dlm)))
End Function
Public Sub WriteTextFile(ByVal filepath As String, ByVal fileText As String)
    Dim fileNum As Integer: fileNum = FreeFile
    If Dir(filepath) <> "" Then Kill filepath
    Open filepath For Output As #fileNum
        Print #fileNum, fileText
    Close #fileNum
End Sub
Public Function COnx(ByRef model As Variant) As IVbOnnx
On Error GoTo err
    Dim obj As Object: Set obj = model
    Set COnx = obj
err:
End Function
