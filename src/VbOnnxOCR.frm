VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxOCR 
   Caption         =   "OnnxOCR"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16965
   OleObjectBlob   =   "VbOnnxOCR.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IVbOnnx
Private WithEvents myglf As GLFrame, selx As Double, sely As Double, flg As Boolean
Attribute myglf.VB_VarHelpID = -1
Private Sub UserForm_Click(): DoEvents: End Sub
Private Property Get IVbOnnx_Editor() As MSForms.TextBox
    Set IVbOnnx_Editor = Me.TextBox3
End Property
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
Public Sub CheckDict(ByRef Parent As VbOnnxMain)
    If flg = False Then
        Dim larr() As Long, i As Long, t As Variant, d
        For Each d In Parent.OnnxResults
            Dim uarr: uarr = d.Item("u_label")
            Dim lstr As String: lstr = ""
            If IsArray(uarr) Then
                ReDim larr(UBound(uarr) - 1) As Long
                For i = LBound(larr) To UBound(larr)
                    larr(i) = uarr(i)
                    If uarr(i) <> "" Then If uarr(i) < 65536 Then lstr = lstr & ChrW(CLng(uarr(i)))
                Next i
            Else
                ReDim larr(0) As Long
            End If
            d.Item("LongArray") = larr
            d.Item("label") = lstr
        Next d
        flg = True
    End If
End Sub
Private Sub IVbOnnx_Render(GLF As GLFrame, Parent As VbOnnxMain)
    Dim ImageAspect As Double, ImageScale As Double
    Dim d As Object, tX As Double, tY As Double, tW As Double, tH As Double, chr() As Byte, idx As Long, text As String
    Dim larr() As Long
    Set myglf = GLF
    ImageScale = GLF.height / (Parent.ImageHeight + 0.1)
    Dim hh As Double: hh = GLF.height * 0.5
    Dim hw As Double: hw = hh * (Parent.ImageWidth / (Parent.ImageHeight + 0.1))
    idx = 0
    CheckDict Parent
    With GLF.gl
        If .UseUnicode = False Then
            Dim chr2() As Byte: chr2 = StrConv("Initializing Unicode fonts...", vbFromUnicode)
            .PushMatrix
                .listbase FONT_BASE_EXLARGE
                .Color4f 1, 1, 1, 1
                .RasterPos3d -0.49 * GLF.width, -0.3 * GLF.height, 2
                .CallLists UBound(chr2) + 1, GL_UNSIGNED_BYTE, VarPtr(chr2(0))
                .listbase 0
                .SwapBuffers
            .PopMatrix
            .UseUnicode = True
            myglf.Refresh
            .SwapBuffers
        Else
            .Enable GL_TEXTURE_2D
                .Begin GL_QUADS
                    .TexCoord2d 0, 1: .Vertex3d -hw, -hh, 1
                    .TexCoord2d 1, 1: .Vertex3d hw, -hh, 1
                    .TexCoord2d 1, 0: .Vertex3d hw, hh, 1
                    .TexCoord2d 0, 0: .Vertex3d -hw, hh, 1
                .End1
            .Disable GL_TEXTURE_2D
            For Each d In Parent.OnnxResults
                With d
                    tX = ImageScale * .Item("x") - hw
                    tY = hh - ImageScale * .Item("y")
                    tW = ImageScale * .Item("w")
                    tH = ImageScale * .Item("h")
                    larr = .Item("LongArray")
                End With
                .LineWidth 2
                .Enable GL_BLEND
                    .BlendFunc GL_SRC_ALPHA, GL_ONE
                    .Color4f 1, 1, 0, 0.75
                    .Begin GL_QUADS
                        .Vertex3d tX, tY, 2
                        .Vertex3d tX + tW, tY, 2
                        .Vertex3d tX + tW, tY - tH, 2
                        .Vertex3d tX, tY - tH, 2
                    .End1
                .Disable GL_BLEND
                .Color4f 1, 0, 0, 0
                .Begin GL_LINE_LOOP
                    .Vertex3d tX, tY, 3
                    .Vertex3d tX + tW, tY, 3
                    .Vertex3d tX + tW, tY - tH, 3
                    .Vertex3d tX, tY - tH, 3
                .End1
                .listbase FONT_BASE_UNICODE
                    .Color4f 0#, 0#, 0.5, 0#
                    .RasterPos3d tX, tY - tH * 0.5, 4
                    .CallLists UBound(larr) + 1, GL_UNSIGNED_INT, VarPtr(larr(0))
                .listbase 0
                idx = idx + 1
            Next d
        End If
    End With
End Sub
Private Function IVbOnnx_Export(target As Worksheet, Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    Dim i As Long, lnColor As Long, asp As Double, tX As Double, tY As Double, tW  As Double, tH As Double, csize, pY As Double
    Dim capAll As String
    Dim iw As Double: iw = Parent.ImageWidth
    Dim ih As Double: ih = Parent.ImageHeight
    asp = iw / (ih + 0.1)
    csize = 300
    CheckDict Parent
    Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * asp, height:=csize)
    With IVbOnnx_Export
        .name = "Scatter" & format(Now(), "yyyymmdd-hhmmss")
        With .Chart
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
                .MaximumScale = ih
                .ReversePlotOrder = True
                .MajorGridlines.format.Line.Visible = msoFalse
                .TickLabels.NumberFormat = ";;;"
                .MajorTickMark = xlTickMarkNone
            End With
            With .Axes.Item(1)
                .MinimumScale = 0
                .MaximumScale = iw
                .TickLabels.NumberFormat = ";;;"
                .MajorTickMark = xlTickMarkNone
            End With
        On Error GoTo err
            For i = 1 To Parent.OnnxResults.Count + 1
                With .SeriesCollection.NewSeries
                    If i <= Parent.OnnxResults.Count Then
                        With Parent.OnnxResults.Item(i)
                            tX = .Item("x")
                            tY = .Item("y")
                            tW = .Item("w")
                            tH = .Item("h")
                            If tX < 0 Then tX = 0
                        End With
                        lnColor = .border.color
                        .name = Parent.OnnxResults.Item(i).Item("label")
                        If Abs(pY - tY) > 20 Then capAll = capAll & vbNewLine & .name Else capAll = capAll & " " & .name
                        capAll = Replace(Replace(Replace(capAll, "  ", " "), vbNewLine & vbNewLine, vbNewLine), vbNewLine & " ", vbNewLine)
                        .ChartType = xlXYScatterLinesNoMarkers
                        .XValues = Array(tX, tX + tW, tX + tW, tX, tX)
                        .Values = Array(tY, tY, tY + tH, tY + tH, tY)
                    Else
                        lnColor = &HFFFFFF
                        .ChartType = xlXYScatterLinesNoMarkers
                        .XValues = Array(0, iw, iw, 0, 0)
                        .Values = Array(0, 0, ih, ih, 0)
                        With .Points(1)
                            .ApplyDataLabels
                            With .DataLabel
                                .AutoText = False
                                .ShowValue = 0
                                .ShowSeriesName = -1
                                .Font.size = 7
                                With .format.TextFrame2
                                    .WordWrap = msoTrue
                                    If capAll = "" Then .TextRange.text = "No data." Else .TextRange.text = capAll
                                End With
                                .HorizontalAlignment = xlLeft
                                .VerticalAlignment = xlTop
                                .Top = 0
                                .Left = csize * asp * 0.35
                                .width = csize * asp * 0.65
                                .height = csize
                            End With
                        End With
                        .LeaderLines.Delete
                    End If
                End With
                pY = tY
            Next i
            With .PlotArea
                .width = csize * asp * 0.35
                .height = csize * 0.35
                .Left = 0
                .Top = 0
            End With
        End With
    End With
err:
End Function
