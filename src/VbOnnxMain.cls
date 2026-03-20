VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbOnnxMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Event FileChecked(CheckedFilePath As String, url As String, Check As Boolean)
Public Event Executed()
Public Event Running()
Public OnnxResults As Collection
Public ImageWidth As Long, ImageHeight As Long, ImagePath As String
Private busy As Boolean, ibuffer() As Byte
Private RequiredFiles As Collection
Private Const DL_INFO As String = _
"The required files is missing." & vbCrLf & _
"Microsoft Edge will now open the download page." & vbCrLf & vbCrLf & _
"Please download the file manually and place it in the Temp folder:"
Private Const DL_INFO2 As String = "Attention : Please use the downloaded files in accordance with the license terms provided by their respective distributors."
Private Const PS_NAME = "server.ps1"
Private Const JS_NAME = "index.html"
Private Const SERVER_NAME = "tmp_server"
Private Const ONNX_LIB = "https://cdn.jsdelivr.net/npm/onnxruntime-web@1.18.0"
Private Const BOX_W = 240
Private Const BOX_H = 90
Private AngleRate As Double
Private SelectedOnnx As IVbOnnx
Private http As Object, sh As Object, glMsg(3) As String
Public ExecuteFlg As Boolean
Public Property Get FILE_FILTER() As String
    FILE_FILTER = "imagefile(*.bmp;*.jpg;*.png),*.bmp;*.jpg;*.png"
End Property
Private Function myModels() As Variant
    myModels = Array(New VbOnnxResNet, New VbOnnxYOLOX, New VbOnnxFNST, New VbOnnxMiDaS, New VbOnnxAnimeGAN)
End Function
Public Function GetModelDict(ParamArray vbonnx()) As Object
    Dim dict, m: Set dict = CreateObject("Scripting.Dictionary")
    If UBound(vbonnx) <= 0 Then vbonnx = myModels
    For Each m In vbonnx
        dict.Add Me.COnx(m).name, m
    Next m
    Set GetModelDict = dict
End Function
Private Function GetPipeLine(ByVal TargetPath As String) As String
    If SelectedOnnx Is Nothing Then Exit Function
    GetPipeLine = "<!doctype html><html>" & vbNewLine & _
                 "<head><meta charset='UTF-8'>" & vbNewLine & _
                    "<title>Running : " & SelectedOnnx.name & "</title>" & vbNewLine & _
                    "<script src='ort.min.js'></script>" & vbNewLine & _
                    "<style>" & vbNewLine & _
                            "body {margin: 0;height: 100vh;display: flex;justify-content: center;align-items: center;background: #f9f9f9;font-family: sans-serif;}" & vbNewLine & _
                            "#loader {border: 8px solid #f3f3f3;border-top: 8px solid #3498db;border-radius: 50%;width: 60px;height: 60px;animation: spin 1s linear infinite;position: absolute;}" & vbNewLine & _
                            "@keyframes spin {0% {transform: rotate(0deg);}100% { transform: rotate(360deg);}}" & vbNewLine & _
                    "</style>" & vbNewLine & _
                 "</head>" & vbNewLine & _
                 "<body>" & vbNewLine & _
                        "<div id='loader'></div><img id='img' src='images/" & GetFileName(TargetPath) & "' style='display:none;' />" & vbNewLine & vbNewLine & _
                        "<script>" & vbNewLine & _
                                "(async () => {" & vbNewLine & _
                                SelectedOnnx.JsCode & vbNewLine & _
                                "})();" & vbNewLine & _
                        "</script>" & vbNewLine & _
                        "</body>" & vbNewLine & _
                 "</html>"
End Function
Public Function Execute(ByVal TargetFilePath As String) As Collection
    If ExecuteFlg Then
        If ImagePath <> TargetFilePath Then
            Call Image2buffer(TargetFilePath)
            FileCopy TargetFilePath, ImagesPath & "\" & GetFileName(TargetFilePath)
            ImagePath = TargetFilePath
        End If
        WriteTextFile JsPath, GetPipeLine(TargetFilePath)
        Set OnnxResults = json2coll(CallRest)
        Set Execute = OnnxResults
        RaiseEvent Executed
    End If
End Function
Private Sub Image2buffer(ByVal TargetFilePath As String)
    Dim imst As ImageStream: Set imst = New ImageStream
    ibuffer = imst.GetBmpStreamFromFile(TargetFilePath)
    ImageWidth = imst.width
    ImageHeight = imst.height
    Set imst = Nothing
End Sub
Private Function CallRest() As String
    Dim result As String
    With http
        .Open "POST", "http://localhost:8080/api/infer", False
        Call Sleep(250)
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        Call Sleep(1000)
        Do
            Call Sleep(250): DoEvents
            RaiseEvent Running
            .Open "GET", "http://localhost:8080/api/getresult", False
            .Send
            result = .ResponseText
        Loop While result = ""
    End With
    CallRest = result
End Function
Private Sub Class_Initialize()
    Call SetRequiredFiles
    AngleRate = 1 / Tan(15 * DEG2RAD)
    If Dir(TempPath, vbDirectory) = "" Then MkDir TempPath
    ExecuteFlg = False
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    glMsg(0) = "Left drag: Move image"
    glMsg(1) = "Right drag (up/down): Zoom"
    glMsg(2) = "Both buttons + drag: Rotate image"
    glMsg(3) = "Double-click: Reset camera"
    Sleep 500
    Call initSever(zPrefs.SeverCode)
End Sub
Private Sub Class_Terminate()
    Call Term
End Sub
Public Sub Term()
On Error GoTo err:
    sh.Terminate: Set http = Nothing: Set sh = Nothing
err:
End Sub
Private Sub initSever(ByVal ServerCode As String)
    Dim pscode As String: pscode = "Set-Location '" & TempPath & "'" & vbNewLine & ServerCode
    If Dir(ImagesPath, vbDirectory) = "" Then MkDir ImagesPath
    WriteTextFile PsPath, pscode
    If sh Is Nothing Then
        Sleep 1000
        Set sh = CreateObject("WScript.Shell").Exec("powershell -NoLogo -ExecutionPolicy Bypass -File " & PsPath)
        Sleep 1000
    End If
End Sub
Public Function COnx(ByRef model As Variant) As IVbOnnx
    Set COnx = model
End Function
Private Sub SetRequiredFiles(Optional ByRef exLibs As Collection = Nothing)
    Dim tlib, stdlibs: Set RequiredFiles = New Collection
    stdlibs = Array("/dist/ort.min.js", "/dist/ort-wasm.wasm", "/dist/ort-wasm-simd.wasm", "/dist/ort-wasm-threaded.wasm", "/dist/ort-wasm-simd-threaded.wasm")
    With RequiredFiles
        For Each tlib In stdlibs
            .Add ONNX_LIB & tlib
        Next tlib
        If Not (exLibs Is Nothing) Then
            For Each tlib In exLibs
                .Add tlib
            Next tlib
        End If
    End With
End Sub
Public Sub OpenTempFolder()
    CreateObject("WScript.Shell").Run TempPath
End Sub
Public Property Get OnnxModel() As IVbOnnx
    Set OnnxModel = SelectedOnnx
End Property
Public Property Set OnnxModel(ByRef tOnnx As IVbOnnx)
    Set SelectedOnnx = tOnnx
    SetRequiredFiles SelectedOnnx.exLibs
    ExecuteFlg = EnsureRuntimeFiles
End Property
Private Property Get PsPath() As String
    PsPath = TempPath & "\" & PS_NAME
End Property
Private Property Get JsPath() As String
    JsPath = TempPath & "\" & JS_NAME
End Property
Public Property Get ImagesPath() As String
    ImagesPath = TempPath & "\images"
End Property
Private Property Get TempPath() As String
On Error GoTo err
    TempPath = (Environ("TEMP")) & "\" & SERVER_NAME
err:
End Property
Public Function EnsureRuntimeFiles() As Boolean
    Dim url As Variant, flg As Boolean: flg = True
    For Each url In RequiredFiles
        flg = flg * EnsureFile(CStr(url), TempPath & "\" & GetFileName(url, "/"))
    Next url
    EnsureRuntimeFiles = flg
End Function
Private Function EnsureFile(ByVal url As String, ByVal localPath As String) As Boolean
    If Dir(localPath) = "" Then
        Call Sleep(250): DoEvents
        On Error GoTo err
        If localPath Like "*.onnx" Then GoTo err
        If localPath Like "*coco*" Then GoTo err
        If localPath Like "*imagenet*" Then GoTo err
        With http
            .Open "GET", url, False
            .Send
            If .status = 200 Then
                With CreateObject("ADODB.Stream")
                    .Type = 1 ' binary
                    .Open
                    .Write http.responseBody
                    .SaveToFile localPath, 2
                    .Close
                End With
            End If
        End With
    End If
    RaiseEvent FileChecked(GetFileName(localPath), url, True)
    EnsureFile = True
    Exit Function
err:
    RaiseEvent FileChecked(GetFileName(localPath), url, False)
    EnsureFile = False
End Function
Public Sub OpenEdgeForDownload()
    Dim url, flg As Boolean, cnt As Long
    cnt = 0
    For Each url In RequiredFiles
        flg = EnsureFile(CStr(url), TempPath & "\" & GetFileName(url, "/"))
        If flg = False Then
            If cnt = 0 Then MsgBox DL_INFO & vbCrLf & TempPath & vbCrLf & vbCrLf & DL_INFO2, vbInformation, "Download ONNX File"
            Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
            wsh.Run "msedge.exe " & url
            cnt = cnt + 1
        End If
    Next url
End Sub
Public Sub ExportModelCode(ByVal TargetFilePath As String)
    If Me.OnnxModel Is Nothing Then Exit Sub
    WriteTextFile TargetFilePath, OnnxModel.JsCode
End Sub
Public Function Export(Optional ByRef target As Worksheet = Nothing, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    If Me.OnnxResults Is Nothing Then Exit Function
    If target Is Nothing Then Set target = ActiveSheet
    Set Export = Me.OnnxModel.Export(target, Me, Left, Top)
End Function
Public Sub Render(ByRef target As GLFrame, Zoom As Double, offsetX As Double, offsetY As Double, pitch As Double, roll As Double, yaw As Double)
    With target
        Dim tHeight As Double, tWidth As Double, AspectRatio As Double, Distance As Double, i As Long
        Dim AoV As Double: AoV = 10
        tHeight = .height
        tWidth = .width
        AspectRatio = tWidth / (tHeight + 0.0001)
        Distance = (0.5 * tHeight * Zoom) / Tan(AoV * 0.5 * DEG2RAD)
        With .gl
            If Not Not ibuffer Then
                .GenTextures 1, 1
                .BindTexture GL_TEXTURE_2D, 1
                .Enable GL_TEXTURE_2D
                    .PixelStorei GL_UNPACK_ALIGNMENT, 1
                    .Build2DMipmaps GL_TEXTURE_2D, GL_RGBA, ImageWidth, ImageHeight, GL_BGRA, GL_UNSIGNED_BYTE, VarPtr(ibuffer(0))
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
                    .TexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
                .Disable GL_TEXTURE_2D
                Erase ibuffer
            End If
            .Viewport 0, 0, tWidth, tHeight
            
            .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
            .MatrixMode GL_PROJECTION
            
            .LoadIdentity
            .Perspective AoV, AspectRatio, Distance * 0.5, Distance * 1.5
                .LookAt -offsetX, -offsetY - 0.001, Distance, -offsetX, -offsetY, 0, 0, 0, 1
            .MatrixMode GL_MODELVIEW
            .BindTexture GL_TEXTURE_2D, 1
            .LoadIdentity
            
            .PushMatrix
                .Rotated pitch, -1, 0, 0
                .Rotated roll, 0, -1, 0
                .Rotated yaw, 0, 0, 1
                .Color4f 1#, 1#, 1#, 1#
                If Not OnnxResults Is Nothing Then
                    If OnnxResults.Count > 0 Then Me.OnnxModel.Render target, OnnxResults, CDbl(ImageWidth / ImageHeight), CDbl(tHeight / ImageHeight)
                End If
            .PopMatrix

            .MatrixMode GL_PROJECTION
            .LoadIdentity
            .Ortho2D 0, tWidth, 0, tHeight
            .MatrixMode GL_MODELVIEW
            .LoadIdentity

            Dim yy As Double
            Dim bx As Double: bx = tWidth - BOX_W - 30
            Dim by As Double: by = 10
            .Color4f 0, 1, 1, 1
            .Begin GL_LINE_LOOP
                .Vertex2d bx, by
                .Vertex2d bx + BOX_W, by
                .Vertex2d bx + BOX_W, by + BOX_H
                .Vertex2d bx, by + BOX_H
            .End1
            yy = by + BOX_H - 20
            .listbase 1000
            For i = 0 To 3
                Dim chr() As Byte: chr = StrConv(glMsg(i), vbFromUnicode)
                .RasterPos2d bx + 10, yy
                .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                yy = yy - 18
            Next i
            .listbase 0
            
            .SwapBuffers
        End With
    End With
End Sub


