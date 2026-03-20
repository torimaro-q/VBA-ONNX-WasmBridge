VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxYOLOX 
   Caption         =   "model"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "VbOnnxYOLOX.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxYOLOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IVbOnnx
Private WithEvents myglf As GLFrame, selx As Double, sely As Double
Attribute myglf.VB_VarHelpID = -1
Private Sub myglf_Click(ByVal X As Double, y As Double, Button As Integer)
    'Write the processing using the result of the hit test (OpenGL) here.
    selx = X
    sely = y
End Sub
Private Property Get IVbOnnx_Name() As String
    IVbOnnx_Name = Me.TextBox1.Value
End Property
Private Property Get IVbOnnx_Info() As String
    IVbOnnx_Info = Me.TextBox2.Value
End Property
Private Property Get IVbOnnx_JsCode() As String
    IVbOnnx_JsCode = Me.TextBox3.Value
End Property
Private Property Get IVbOnnx_exLibs() As Collection
    Dim arr As Variant: arr = Split(Me.TextBox4.Value, vbNewLine)
    Dim tmp, coll As Collection: Set coll = New Collection
    For Each tmp In arr
        coll.Add Application.Clean(tmp)
    Next tmp
    Set IVbOnnx_exLibs = coll
End Property
Private Sub IVbOnnx_Render(GLF As GLFrame, Results As Collection, Optional ByVal imageAspect As Double = 1#, Optional ByVal imageScale As Double = 1#)
    Set myglf = GLF
    Dim d As Object, zofs As Double, tX, tY, tw, th, score, chr() As Byte, idx As Long
    Dim hh As Double: hh = GLF.height * 0.5
    Dim hw As Double: hw = hh * imageAspect
    idx = 0
    With GLF.gl
        .Enable GL_TEXTURE_2D
            .Begin GL_QUADS
                .TexCoord2d 0, 1: .Vertex3d -hw, -hh, 1
                .TexCoord2d 1, 1: .Vertex3d hw, -hh, 1
                .TexCoord2d 1, 0: .Vertex3d hw, hh, 1
                .TexCoord2d 0, 0: .Vertex3d -hw, hh, 1
            .End1
        .Disable GL_TEXTURE_2D
        
        For Each d In Results
                With d
                    tX = imageScale * .Item("x") - hw
                    tY = hh - imageScale * .Item("y")
                    tw = imageScale * .Item("w")
                    th = imageScale * .Item("h")
                    score = format(.Item("score"), "0%")
                    chr = StrConv(d.Item("label") & ": " & score, vbFromUnicode)
                End With
                
                .Enable GL_BLEND
                .BlendFunc GL_SRC_ALPHA, GL_ONE
                .Color4f 0.5, 0.5, 0#, 0.4
                
                .Begin GL_QUADS
                    .Vertex3d tX, tY, 10 + idx
                    .Vertex3d tX + tw, tY, 10 + idx
                    .Vertex3d tX + tw, tY - th, 10 + idx
                    .Vertex3d tX, tY - th, 10 + idx
                .End1
                
                .LineWidth 2
                .Color4f 1, 1, 1, 0.8
                .Begin GL_LINE_LOOP
                    .Vertex3d tX, tY, 20 + idx
                    .Vertex3d tX + tw, tY, 20 + idx
                    .Vertex3d tX + tw, tY - th, 20 + idx
                    .Vertex3d tX, tY - th, 20 + idx
                .End1
                
                .Disable GL_BLEND
                zofs = 0
                
                .listbase 2000
                .Color4f 1#, 1#, 1#, 1#
                .RasterPos3d tX, tY, 40 + idx + zofs
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                
                .Color4f 0#, 0#, 0#, 0#
                .RasterPos3d tX + 3, tY - 3, 30 + idx + zofs
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                .listbase 0
                
                idx = idx + 1
            Next d
    End With
End Sub
Private Function IVbOnnx_Export(target As Worksheet, Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    Dim i As Long, lnColor As Long, asp As Double, tX, tY, w, h, csize
    asp = Parent.ImageWidth / (Parent.ImageHeight + 0.1)
    csize = 300
    Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * asp, height:=csize)
    With IVbOnnx_Export
        .name = "Scatter" & format(Now(), "yyyymmdd-hhmmss")
        With .Chart
            .ChartType = xlXYScatterLines
            .HasLegend = True
            .Legend.Delete
            With .PlotArea.format.Fill
                .Visible = msoTrue
                If Parent.ImagePath <> "" Then
                    .UserPicture Parent.ImagePath
                    .TextureTile = msoFalse
                End If
            End With
            With .Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = Parent.ImageHeight
                .ReversePlotOrder = True
                .MajorGridlines.format.Line.Visible = msoFalse
            End With
            With .Axes
                .Item(1).MinimumScale = 0
                .Item(1).MaximumScale = Parent.ImageWidth
            End With
        On Error GoTo err
            For i = 1 To Parent.OnnxResults.Count
                With .SeriesCollection.NewSeries
                    lnColor = .border.Color
                    .name = Parent.OnnxResults.Item(i).Item("label") & ":" & format(Parent.OnnxResults.Item(i).Item("score"), "0.0%")
                    .ChartType = xlXYScatterLines
                    With Parent.OnnxResults.Item(i)
                        tX = .Item("x")
                        If tX < 0 Then tX = 0
                        tY = .Item("y")
                        w = .Item("w")
                        h = .Item("h")
                    End With
                    .XValues = Array(tX, tX + w, tX + w, tX, tX)
                    .Values = Array(tY, tY, tY + h, tY + h, tY)
                    With .Points(1)
                        .ApplyDataLabels
                        With .DataLabel
                            With .format.Fill
                                .Visible = msoTrue
                                .ForeColor.RGB = RGB(255, 255, 255)
                                .Transparency = 0
                                .Solid
                            End With
                            With .format.Line
                                .Visible = msoTrue
                                .ForeColor.RGB = lnColor
                                .Transparency = 0
                                .Visible = msoTrue
                                .weight = 1.5
                            End With
                            .ShowValue = 0
                            .ShowSeriesName = -1
                        End With
                    End With
                End With
            Next i
        End With
    End With
err:
End Function
