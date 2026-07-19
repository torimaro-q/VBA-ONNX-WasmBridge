VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxUF_FER 
   Caption         =   "UF_FER"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16965
   OleObjectBlob   =   "VbOnnxUF_FER.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxUF_FER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IVbOnnx
Private WithEvents myglf As GLFrame, flg As Boolean, emotions As Variant, rx As Variant, ry As Variant
Attribute myglf.VB_VarHelpID = -1
Private Property Get IVbOnnx_Editor() As MSForms.TextBox
    Set IVbOnnx_Editor = Me.TextBox3
End Property
Private Sub UserForm_Click(): DoEvents: End Sub
Private Property Get IVbOnnx_Name() As String
    IVbOnnx_Name = Me.Caption
End Property
Private Property Get IVbOnnx_Info() As String
    IVbOnnx_Info = TextBox1.Value
End Property
Private Property Get IVbOnnx_JsCode() As String
    IVbOnnx_JsCode = TextBox3.Value
    flg = False
End Property
Private Property Get IVbOnnx_exLibs() As Collection
    Dim arr As Variant: arr = Split(TextBox2.Value, vbNewLine)
    Dim tmp, coll As Collection: Set coll = New Collection
    For Each tmp In arr
        coll.Add Application.Clean(tmp)
    Next tmp
    Set IVbOnnx_exLibs = coll
End Property
Private Sub UserForm_Initialize()
    emotions = Array("neutral", "happiness", "surprise", "sadness", "anger", "disgust", "fear", "contempt")
    ReDim rx(UBound(emotions))
    ReDim ry(UBound(emotions))
    Dim i As Long, dth As Double: dth = 6.28 / (UBound(emotions) + 1)
    For i = 0 To UBound(emotions)
        rx(i) = 100 * Cos(i * dth)
        ry(i) = 100 * Sin(i * dth)
    Next i
End Sub
Private Sub IVbOnnx_Render(GLF As GLFrame, Parent As VbOnnxMain)
    Dim ImageScale As Double, results As Collection
    ImageScale = GLF.height / (Parent.ImageHeight + 0.1)
    Set results = Parent.OnnxResults
    Set myglf = GLF
    Dim hh As Double: hh = GLF.height * 0.5
    Dim hw As Double: hw = hh * (Parent.ImageWidth / (Parent.ImageHeight + 0.1))
    Dim d As Object, zofs As Double, tX, tY, tW, tH, score, chr() As Byte, idx As Long
    Dim i As Long, t As Double
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
        For Each d In results
            With d
                tX = ImageScale * .Item("x") - hw
                tY = hh - ImageScale * .Item("y")
                tW = ImageScale * .Item("w")
                tH = ImageScale * .Item("h")
                score = format(.Item("eScore"), "0%")
                chr = StrConv(d.Item("emotion") & ": " & score, vbFromUnicode)
            End With
            .Enable GL_BLEND
            .BlendFunc GL_SRC_ALPHA, GL_ONE
            .Color4f 0.5, 0.5, 0#, 0.1
            .Begin GL_QUADS
                .Vertex3d tX, tY, 2 + idx
                .Vertex3d tX + tW, tY, 2 + idx
                .Vertex3d tX + tW, tY - tH, 2 + idx
                .Vertex3d tX, tY - tH, 2 + idx
            .End1
            .LineWidth 2
            .Color4f 1, 1, 1, 0.3
            .Begin GL_LINE_LOOP
                .Vertex3d tX, tY, 22 + idx
                .Vertex3d tX + tW, tY, 22 + idx
                .Vertex3d tX + tW, tY - tH, 22 + idx
                .Vertex3d tX, tY - tH, 22 + idx
            .End1
            .Color4f 1#, 1#, 1#, 1#
            .Begin GL_LINE_LOOP
                For i = 0 To UBound(emotions)
                    .Vertex3d rx(i) + tX - 15, ry(i) + tY - 15, 30
                Next i
            .End1
            .Color4f 0#, 0.5, 1, 0.4
            .Begin GL_POLYGON
                For i = 0 To UBound(emotions)
                    t = d.Item(emotions(i))
                    .Vertex3d t * rx(i) + tX - 15, t * ry(i) + tY - 15, 30
                Next i
            .End1
            .Color4f 0#, 0.5, 1, 0.8
            .Begin GL_LINE_LOOP
                For i = 0 To UBound(emotions)
                    t = d.Item(emotions(i))
                    .Vertex3d t * rx(i) + tX - 15, t * ry(i) + tY - 15, 30
                Next i
            .End1
            .Disable GL_BLEND
            .listbase FONT_BASE_SMALL
            For i = 0 To UBound(emotions)
                chr = StrConv(emotions(i) & " : " & format(d.Item(emotions(i)), "0%"), vbFromUnicode)
                .Color4f 1#, 1#, 1#, 1#
                .RasterPos3d rx(i) + tX - 15, ry(i) + tY - 15, 50
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
            Next i
            .listbase 0
            idx = idx + 1
        Next d
    End With
End Sub
Private Function IVbOnnx_Export(target As Worksheet, Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    Dim i As Long, j As Long, lnColor As Long, asp As Double, tX As Double, tY As Double, W, h, csize
    asp = Parent.ImageWidth / (Parent.ImageHeight + 0.1)
    csize = 300
    Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * asp, height:=csize)
    With IVbOnnx_Export
        .name = "Scatter" & format(Now(), "yyyymmdd-hhmmss")
        With .Chart
            If Parent.OnnxResults.Count > 1 Then
                .ChartType = xlXYScatterLinesNoMarkers
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
                        lnColor = .border.color
                        .name = Parent.OnnxResults.Item(i).Item("emotion") & ":" & format(Parent.OnnxResults.Item(i).Item("eScore"), "0.0%")
                        .ChartType = xlXYScatterLinesNoMarkers
                        With Parent.OnnxResults.Item(i)
                            tX = .Item("x")
                            tY = .Item("y")
                            W = .Item("w")
                            h = .Item("h")
                            If tX < 0 Then tX = 0
                        End With
                        .XValues = Array(tX, tX + W, tX + W, tX, tX)
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
            Else
            
            On Error GoTo err
                Dim tResult As Variant: Set tResult = Parent.OnnxResults.Item(1)
                Dim emlabel As Variant: ReDim emlabel(LBound(emotions) To UBound(emotions))
                Dim emscore As Variant: ReDim emscore(LBound(emotions) To UBound(emotions))
                For i = LBound(emotions) To UBound(emotions)
                    emlabel(i) = emotions(i) & vbCr & format(tResult.Item(emotions(i)), "0%")
                    emscore(i) = tResult.Item(emotions(i))
                Next i
                If Rnd() > 0.5 Then
                    .ChartType = xlPie
                    With .SeriesCollection.NewSeries
                        .name = tResult.Item("emotion")
                        .XValues = emlabel
                        .Values = emscore
                    End With
                    With .PlotArea
                        .width = csize * asp * 0.45
                        .Top = csize * 0.2
                    End With
                    .HasLegend = True
                    With .Legend
                        .Fill.UserPicture Parent.ImagePath
                        .Left = csize * asp * 0.6
                        .width = csize * asp * 0.55
                        .Top = csize * 0.2
                        .height = (.width / asp)
                        If .LegendEntries.Count > 2 Then
                            For i = .LegendEntries.Count To 2 Step -1
                                .LegendEntries(i).Delete
                            Next i
                            With .format.TextFrame2.TextRange.Font
                                .size = 20
                                .Bold = msoTrue
                                .Caps = msoNoCaps
                                With .Fill
                                    .Visible = msoTrue
                                    .ForeColor.RGB = RGB(254, 254, 254)
                                    .Transparency = 0
                                    .Solid
                                End With
                                With .Line
                                    .Visible = msoTrue
                                    .Transparency = 0
                                    .weight = 0.75
                                    .DashStyle = msoLineSolid
                                    .style = msoLineSingle
                                    With .ForeColor
                                        .ObjectThemeColor = msoThemeColorAccent1
                                        .TintAndShade = 0
                                        .Brightness = 0
                                    End With
                                End With
                                .Spacing = 0.5
                            End With
                        End If
                    End With
                Else
                    With .ChartArea.Fill
                        .Visible = msoTrue
                        .UserPicture Parent.ImagePath
                    End With
                    .ChartType = xlRadarFilled
                    .HasLegend = False
                    With .Axes(xlValue)
                        .MinimumScale = 0
                        .MaximumScale = 1
                        .TickLabels.NumberFormatLocal = "0%"
                        .format.Line.Visible = msoFalse
                        .MajorUnit = 1
                        .Delete
                    End With
                    With .SeriesCollection.NewSeries
                        .name = tResult.Item("emotion")
                        .XValues = emlabel
                        .Values = emscore
                    End With
                    .ChartTitle.Delete
                    .PlotArea.Fill.Visible = msoFalse
                    With .ChartArea.format.TextFrame2.TextRange.Font
                        .size = 12
                        .Bold = msoTrue
                        .Caps = msoNoCaps
                        With .Fill
                            .Visible = msoTrue
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Transparency = 0
                            .Solid
                        End With
                        With .Line
                            .Visible = msoTrue
                            .Transparency = 0
                            .weight = 0.5
                            .DashStyle = msoLineSolid
                            .style = msoLineSingle
                            .ForeColor.RGB = RGB(254, 254, 254)
                        End With
                    End With
                End If
            End If
        End With
    End With
err:
End Function
