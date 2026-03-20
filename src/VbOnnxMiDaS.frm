VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxMiDaS 
   Caption         =   "model"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "VbOnnxMiDaS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxMiDaS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IVbOnnx
Private Const CFF As Double = 1 / 256
Private vecs() As Vector3d
Private texs() As Vector2d
Private Property Get IVbOnnx_Name() As String
    IVbOnnx_Name = Me.TextBox1.Value
End Property
Private Property Get IVbOnnx_Info() As String
    IVbOnnx_Info = Me.TextBox2.Value
End Property
Private Property Get IVbOnnx_JsCode() As String
    IVbOnnx_JsCode = Me.TextBox3.Value
    ReDim vecs(0) As Vector3d
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
    Dim i As Long, lnColor As Long, asp As Double, j As Long
    Dim tX, tY, w, h, name
    asp = Parent.ImageWidth / (Parent.ImageHeight + 0.1)
    name = "Scatter" & format(Now(), "yyyymmdd-hhmmss")
    Dim csize
    csize = 300
    Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * 1.2, height:=csize)
    With IVbOnnx_Export
        .name = name
        With .Chart
            .ChartColor = 18
            .HasLegend = True
            Dim depth: depth = Parent.OnnxResults.Item(1).Item("depth")
            For i = 0 To 255
                If i = 255 Then Exit For
                With .SeriesCollection.NewSeries
                    lnColor = .border.Color
                    .name = CStr(i)
                    Dim arr(): ReDim arr(255)
                    Dim dep(): ReDim dep(255)
                    For j = 0 To 255
                        arr(j) = j
                        dep(j) = CDbl(depth(i * 256 + j))
                    Next j
                    .XValues = arr
                    .Values = dep
                    If i > 1 Then .ChartType = xlSurface
                End With
            Next i
            .Axes(xlSeries).ReversePlotOrder = True
            With .ChartArea.format.ThreeD
                .RotationX = -10
                .RotationY = 170
            End With
            .Axes(xlValue).MajorUnit = 10
        End With
    End With
End Function
Private Sub IVbOnnx_Render(GLF As GLFrame, Results As Collection, Optional ByVal imageAspect As Double = 1#, Optional ByVal imageScale As Double = 1#)
    Dim depthA As Variant, d As Object
    Dim i As Long, h As Long, w As Long
    Dim hw1 As Long, hw2 As Long, hh1 As Long, hh2 As Long
    Dim dw As Double: dw = (GLF.height * imageAspect) * CFF
    Dim dh As Double: dh = (GLF.height) * CFF
    Dim size As Long
    For Each d In Results
        With d
            If .Exists("depth") Then
                If UBound(vecs) < 10 Then
                    size = (256 ^ 2) * 4 - 1
                    ReDim vecs(size) As Vector3d
                    ReDim texs(size) As Vector2d
                    depthA = .Item("depth")
                    i = 0
                    For h = 0 To -255 Step -1
                        hh1 = dh * (h + 127)
                        hh2 = dh * (h + 128)
                        For w = 0 To 255 Step 1
                            hw1 = dw * (w - 128)
                            hw2 = dw * (w - 127)
                            texs(i + 0) = Vector2d(CFF * w, -CFF * (h - 1))
                            texs(i + 1) = Vector2d(CFF * (w + 1), -CFF * (h - 1))
                            texs(i + 2) = Vector2d(CFF * (w + 1), -CFF * (h))
                            texs(i + 3) = Vector2d(CFF * w, -CFF * (h))
                            If (w = 255) Or (h = -255) Then
                                vecs(i + 0) = Vector3d(hw1, hh1, CDbl(depthA((-h) * 256 + w)))
                                vecs(i + 1) = Vector3d(hw2, hh1, CDbl(depthA((-h) * 256 + w)))
                                vecs(i + 2) = Vector3d(hw2, hh2, CDbl(depthA(-h * 256 + w)))
                                vecs(i + 3) = Vector3d(hw1, hh2, CDbl(depthA(-h * 256 + w)))
                            Else
                                vecs(i + 0) = Vector3d(hw1, hh1, CDbl(depthA((-h + 1) * 256 + w)))
                                vecs(i + 1) = Vector3d(hw2, hh1, CDbl(depthA((-h + 1) * 256 + w + 1)))
                                vecs(i + 2) = Vector3d(hw2, hh2, CDbl(depthA(-h * 256 + w + 1)))
                                vecs(i + 3) = Vector3d(hw1, hh2, CDbl(depthA(-h * 256 + w)))
                            End If
                            i = i + 4
                        Next w
                    Next h
                End If
            End If
        End With
    Next d
    With GLF.gl
        .PushMatrix
        .Enable GL_TEXTURE_2D
            .EnableClientState GL_TEXTURE_COORD_ARRAY
            .EnableClientState GL_VERTEX_ARRAY
                .TexCoordPointer 2, GL_DOUBLE, 0, VarPtr(texs(0))
                .VertexPointer 3, GL_DOUBLE, 0, VarPtr(vecs(0))
                .DrawArrays GL_QUADS, 0, UBound(vecs) + 1
            .DisableClientState GL_VERTEX_ARRAY
            .DisableClientState GL_TEXTURE_COORD_ARRAY
        .Disable GL_TEXTURE_2D
        .PopMatrix
    End With
End Sub

Private Sub UserForm_Click()

End Sub
