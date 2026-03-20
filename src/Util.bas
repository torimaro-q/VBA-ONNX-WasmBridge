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
    b As Byte
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
    V As Long
End Type
Public Const ENCODER_BMP    As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Public Const ENCODER_JPG    As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Public Const ENCODER_PNG    As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Public Const PICTYPE_BITMAP = 1
Public Const CLMASK As Long = &H80000000
Public Const LOGPIXELSX As Long = 88
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const DEG2RAD As Double = 0.017453293

Public Function json2coll(ByVal json As String) As Collection
    Dim buf As Variant: buf = Split(Replace(Replace(json, "]", ""), "[", ""), "{")
    Dim coll As Collection: Set coll = New Collection
    Dim b As Variant
    For Each b In buf
        If Len(b) > 1 Then
            Dim elms, e, k, V, spl
            Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
            With dic
                elms = Split(Replace(b, "}", ""), ",""")
                For Each e In elms
                    If Len(e) > 1 Then
                        spl = Split(e, ":")
                        k = Replace(CStr(spl(0)), """", "")
                        V = Replace(CStr(spl(1)), """", "")
                        If IsNumeric(V) Then
                            .Add k, CDbl(V)
                        Else
                            If V Like "*,*" Then
                                .Add k, Split(V, ",")
                            Else
                                .Add k, V
                            End If
                        End If
                    End If
                Next e
            End With
            coll.Add dic
        End If
    Next b
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
