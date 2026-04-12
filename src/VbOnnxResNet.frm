VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbOnnxResNet 
   Caption         =   "model"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "VbOnnxResNet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "VbOnnxResNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IVbOnnx
Private Sub UserForm_Click(): DoEvents: End Sub
Private Property Get IVbOnnx_Name() As String
    IVbOnnx_Name = TextBox1.Value
End Property
Private Property Get IVbOnnx_Info() As String
    IVbOnnx_Info = TextBox2.Value
End Property
Private Property Get IVbOnnx_JsCode() As String
    IVbOnnx_JsCode = TextBox3.Value
End Property
Private Property Get IVbOnnx_exLibs() As Collection
    Dim arr As Variant: arr = Split(TextBox4.Value, vbNewLine)
    Dim tmp, coll As Collection: Set coll = New Collection
    For Each tmp In arr
        coll.Add Application.Clean(tmp)
    Next tmp
    Set IVbOnnx_exLibs = coll
End Property
Private Sub IVbOnnx_Render(GLF As GLFrame, Parent As VbOnnxMain)
    Dim d As Object, score, chr() As Byte, idx As Long
    Dim hh As Double: hh = GLF.height * 0.5
    Dim hw As Double: hw = hh * (Parent.ImageWidth / (Parent.ImageHeight + 0.1))
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
        For Each d In Parent.OnnxResults
            idx = idx + 1
            With d
                score = d.Item("probability")
                chr = StrConv(d.Item("label") & ": " & format(score, "0.0%"), vbFromUnicode)
            End With
            If score > 0.3 Then
                .listbase FONT_BASE_LARGE
                .Color4f 1#, 1#, 1#, 1#
                .RasterPos3d -hw, idx * 20, 10
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                .Color4f 0#, 0#, 0#, 1#
                .RasterPos3d -hw + 2, idx * 20 - 2, 8
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                .listbase 0
            End If
        Next d
    End With
End Sub
Private Function IVbOnnx_Export(target As Worksheet, Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    Dim i As Long, lnColor As Long, asp As Double, tX
    asp = Parent.ImageWidth / (Parent.ImageHeight + 0.1)
    Dim csize: csize = 300
    Set IVbOnnx_Export = target.ChartObjects.Add(Left:=Left, Top:=Top, width:=csize * asp, height:=csize)
    With IVbOnnx_Export
        .name = "bar" & format(Now(), "yyyymmdd-hhmmss")
        With .Chart
            .ChartType = xlColumnClustered
            .HasLegend = True
            With .Legend
                .Fill.UserPicture Parent.ImagePath
                .Left = csize * asp * 0.4
                .width = csize * asp * 0.6
                .Top = csize * 0.05
                .height = (.width / asp)
            End With
            With .PlotArea
                .width = csize * asp * 0.5
            End With
            With .Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 1
                .MajorGridlines.format.Line.Visible = msoFalse
            End With
        On Error GoTo err
            For i = 1 To Parent.OnnxResults.Count
                With .SeriesCollection.NewSeries
                    lnColor = .border.Color
                    .name = Parent.OnnxResults.Item(i).Item("label") & ":" & format(Parent.OnnxResults.Item(i).Item("probability"), "0.0%")
                    .ChartType = xlColumnClustered
                    With Parent.OnnxResults.Item(i)
                        tX = .Item("probability")
                    End With
                    .XValues = Array("probability")
                    .Values = Array(tX)
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
            .ChartGroups(1).Overlap = -25
            With .Legend
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
        End With
    End With
err:
End Function
