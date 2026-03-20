VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxFNST 
   Caption         =   "model"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "VbOnnxFNST.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxFNST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IVbOnnx
Private b() As Byte, fpath As String, iw As Long, ih As Long
Private Property Get IVbOnnx_Name() As String
    IVbOnnx_Name = Me.TextBox1.Value
End Property
Private Property Get IVbOnnx_Info() As String
    IVbOnnx_Info = Me.TextBox2.Value
End Property
Private Property Get IVbOnnx_JsCode() As String
    IVbOnnx_JsCode = Me.TextBox3.Value
    ReDim b(0) As Byte
End Property
Private Property Get IVbOnnx_exLibs() As Collection
    Dim arr As Variant: arr = Split(Me.TextBox4.Value, vbNewLine)
    Dim tmp, coll As Collection: Set coll = New Collection
    For Each tmp In arr
        coll.Add Application.Clean(tmp)
    Next tmp
    Set IVbOnnx_exLibs = coll
End Property
Private Function IVbOnnx_Export(target As Worksheet, Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    Dim lnColor As Long, asp As Double
    Dim csize: csize = 300
    asp = Parent.ImageWidth / (Parent.ImageHeight + 0.001)
    If CheckBuffer(Parent.OnnxResults) Then
        fpath = Parent.ImagesPath & "\fnst.bmp"
        Dim img As ImageStream: Set img = New ImageStream
        Call img.SaveAsFile(b, iw, ih, fpath)
        Dim name: name = "FNST" & format(Now(), "yyyymmdd-hhmmss")
        Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * asp, height:=csize)
        With IVbOnnx_Export
            .name = name
        End With
        With target.Shapes(name).Fill
            .Visible = msoTrue
            .UserPicture fpath
            .TextureTile = msoFalse
        End With
    End If
End Function
Private Sub IVbOnnx_Render(GLF As GLFrame, Results As Collection, Optional ByVal imageAspect As Double = 1#, Optional ByVal imageScale As Double = 1#)
    Dim hh As Double: hh = GLF.height * 0.5
    Dim hw As Double: hw = hh * imageAspect
    If CheckBuffer(Results) Then
        With GLF
            With .gl
                .GenTextures 1, 2
                .BindTexture GL_TEXTURE_2D, 2
                .Enable GL_TEXTURE_2D
                    .PixelStorei GL_UNPACK_ALIGNMENT, 1
                    .Build2DMipmaps GL_TEXTURE_2D, GL_RGBA, iw, ih, GL_BGRA, GL_UNSIGNED_BYTE, VarPtr(b(0))
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
                    .Begin GL_QUADS
                        .TexCoord2d 0, 1: .Vertex3d -hw, -hh, 1
                        .TexCoord2d 1, 1: .Vertex3d hw, -hh, 1
                        .TexCoord2d 1, 0: .Vertex3d hw, hh, 1
                        .TexCoord2d 0, 0: .Vertex3d -hw, hh, 1
                    .End1
                .Disable GL_TEXTURE_2D
                .BindTexture GL_TEXTURE_2D, 1
            End With
        End With
    End If
End Sub
Private Function CheckBuffer(ByRef Results As Collection) As Boolean
On Error GoTo err
    Dim arrRaw, bsize As Long, i As Long, j As Long
    If UBound(b) < 10 Then
        With Results.Item(1)
            If .Exists("array") Then
                arrRaw = .Item("array")
                iw = .Item("width")
                ih = .Item("height")
                bsize = 4 * iw * ih - 1
                ReDim b(bsize) As Byte
                j = 0
                For i = 0 To bsize Step 4
                    b(i + 0) = arrRaw(j)
                    b(i + 1) = arrRaw(j + 1)
                    b(i + 2) = arrRaw(j + 2)
                    b(i + 3) = 0
                    j = j + 3
                Next i
                CheckBuffer = True
                Exit Function
            Else
                CheckBuffer = False
            End If
        End With
    Else
        CheckBuffer = True
        Exit Function
    End If
err:
    CheckBuffer = False
End Function

Private Sub UserForm_Click()

End Sub
