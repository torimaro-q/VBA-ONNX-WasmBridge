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
Private Const ORT_NAMES = "ort.min.js,ort-wasm-threaded.wasm,ort-wasm.jsep.wasm,ort-wasm-simd-threaded.jsep.mjs,ort-wasm-simd-threaded.jsep.wasm,ort-wasm-simd-threaded.jsep.worker.js,ort-wasm-simd.jsep.wasm"
Private Const ONNX_LIB = "https://cdn.jsdelivr.net/npm/onnxruntime-web@1.23.2/dist/"
Private Const ONNX_CSS = "body {margin: 0;height: 100vh;display: flex;justify-content: center;align-items: center;background: #f9f9f9;font-family: sans-serif;}" & _
                         "#loader {border: 8px solid #f3f3f3;border-top: 8px solid #3498db;border-radius: 50%;width: 60px;height: 60px;animation: spin 1s linear infinite;position: absolute;}" & _
                         "@keyframes spin {0% {transform: rotate(0deg);}100% { transform: rotate(360deg);}}"
Private Const DL_INFO1 As String = "The required files is missing." & vbCrLf & "Microsoft Edge will now open the download page." & vbCrLf & vbCrLf & "Please download the file manually and place it in the Temp folder:"
Private Const DL_INFO2 As String = "Attention : Please use the downloaded files in accordance with the license terms provided by their respective distributors."
Private Const GL_INFO As String = "Left drag: Move image,Right drag (up/down): Zoom,Both buttons + drag: Rotate image,Double-click: Reset camera"
Private Const API_URL = "http://localhost:8080/api/"
Private Const PS_NAME = "server.ps1", JS_NAME = "index.html", SERVER_NAME = "tmp_server", BW = 240, BH = 90
Public Event FileChecked(CheckedPath As String, url As String, Check As Boolean)
Public Event Executed()
Public Event Running()
Public ModelDict As Object, OnnxResults As Collection, ImageWidth As Long, ImageHeight As Long, ImagePath As String, ExecuteFlg As Boolean
Private glMsg As Variant, busy As Boolean, ibuffer() As Byte
Private rtFiles As Collection, exFiles As Collection, sOnnx As IVbOnnx, http As Object, sh As Object, LCache As Object, MCache As Object
Public Property Get FILE_FILTER() As String
    FILE_FILTER = "imagefile(*.bmp;*.jpg;*.png),*.bmp;*.jpg;*.png"
End Property
Private Sub GenModelDict()
On Error GoTo err:
    Dim n, arr: arr = ModelNames
    Set ModelDict = CreateObject("Scripting.Dictionary")
    Set MCache = CreateObject("Scripting.Dictionary")
    For Each n In arr
        ModelDict.Add Replace(n, "VbOnnx", ""), n
    Next n
err:
End Sub
Public Function GetFormFromName(ByVal name As String) As Object
    With ModelDict
        If .exists(name) Then Set GetFormFromName = UserForms.Add(.Item(name))
    End With
End Function
Public Function GetModelFromName(ByVal name As String) As IVbOnnx
    If ModelDict.exists(name) Then
        With MCache
            If .exists(name) = False Then .Add name, UserForms.Add(ModelDict.Item(name))
            Set GetModelFromName = .Item(name)
        End With
    End If
End Function
Private Function GetPipeLine(ByVal tPath As String) As String
    If sOnnx Is Nothing Then Exit Function
    With sOnnx
        GetPipeLine = "<!doctype html><html><head><meta charset='UTF-8'><title>Running : " & .name & "</title><script src='ort.min.js'></script>" & _
                           "<style>" & ONNX_CSS & "</style></head>" & _
                           "<body>" & _
                                  "<div id='loader'></div><img id='img' src='images/" & GetFileName(tPath) & "' style='display:none;' />" & _
                                  "<script>(async () => {" & .JsCode & "})();</script>" & _
                           "</body>" & _
                        "</html>"
    End With
End Function
Public Function Execute(ByVal tPath As String) As Collection
    If ExecuteFlg Then
        If ImagePath <> tPath Then
            Call Image2buffer(tPath)
            FileCopy tPath, ImagesPath & "\" & GetFileName(tPath)
            ImagePath = tPath
        End If
        WriteTextFile JsPath, GetPipeLine(tPath)
        Set OnnxResults = json2coll(CallRest)
        Set Execute = OnnxResults
        RaiseEvent Executed
    End If
End Function
Private Sub Image2buffer(ByVal tPath As String)
    Dim imst As ImageStream: Set imst = New ImageStream
    With imst
        ibuffer = .GetBmpStreamFromFile(tPath)
        ImageWidth = .width
        ImageHeight = .height
    End With
    Set imst = Nothing
End Sub
Private Function CallRest() As String
    Dim result() As Byte, rstr, t
    With http
        .Open "POST", API_URL & "infer", False
        Call Sleep(1000)
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .Send
        Call Sleep(1000)
        t = Timer
        Do
            If Abs(t - Timer) > 60 Then
                .Open "POST", API_URL & "result", False
                .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
                .Send "{""error"":""timeout""}"
            End If
            Call Sleep(1000): DoEvents
            RaiseEvent Running
            .Open "GET", API_URL & "getresult", False
            .Send
            rstr = .ResponseBody
        Loop While Len(rstr) = 0
        result = rstr
        With CreateObject("ADODB.Stream")
            .Type = 1
            .Open
            .Write result
            .Position = 0
            .Type = 2
            .Charset = "utf-8"
            CallRest = .ReadText
            .Close
        End With
    End With
End Function
Private Sub Class_Initialize()
    Call GenModelDict
    Set LCache = CreateObject("Scripting.Dictionary")
    Call SetRequiredFiles
    If Dir(TempPath, vbDirectory) = "" Then MkDir TempPath
    ExecuteFlg = False
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    glMsg = Split(GL_INFO, ",")
    Dim tlib As Variant
    Set rtFiles = New Collection
    With rtFiles
        For Each tlib In Split(ORT_NAMES, ",")
            .Add ONNX_LIB & tlib
        Next tlib
    End With
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
    End If
End Sub
Private Sub SetRequiredFiles(Optional ByRef exLibs As Collection = Nothing)
    Dim tlib: Set exFiles = New Collection
    With exFiles
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
    Set OnnxModel = sOnnx
End Property
Public Property Set OnnxModel(ByRef tOnnx As IVbOnnx)
    Set sOnnx = tOnnx
    SetRequiredFiles sOnnx.exLibs
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
    For Each url In rtFiles
        flg = flg * EnsureFile(CStr(url), TempPath & "\" & GetFileName(url, "/"), True)
    Next url
    For Each url In exFiles
        flg = flg * EnsureFile(CStr(url), TempPath & "\" & GetFileName(url, "/"), False)
    Next url
    ExecuteFlg = flg
    EnsureRuntimeFiles = flg
End Function
Private Function EnsureFile(ByVal url As String, ByVal localPath As String, Optional AutoDownload = True) As Boolean
    If LCache.exists(url) = False Then
        If Dir(localPath) = "" Then
            DoEvents
            On Error GoTo err
            If AutoDownload = False Then GoTo err
            With http
                .Open "GET", url, False
                .Send
                If .status = 200 Then
                    With CreateObject("ADODB.Stream")
                        .Type = 1
                        .Open
                        .Write .ResponseBody
                        .SaveToFile localPath, 2
                        .Close
                    End With
                End If
            End With
        End If
        Call LCache.Add(url, GetFileName(localPath))
        RaiseEvent FileChecked(GetFileName(localPath), url, True)
    Else
        RaiseEvent FileChecked(LCache.Item(url), url, True)
    End If
    EnsureFile = True
    Exit Function
err:
    RaiseEvent FileChecked(GetFileName(localPath), url, False)
    EnsureFile = False
End Function
Public Sub OpenEdgeForDownload()
    Dim url, flg As Boolean, cnt As Long: cnt = 0
    For Each url In exFiles
        flg = EnsureFile(CStr(url), TempPath & "\" & GetFileName(url, "/"), False)
        If flg = False Then
            If cnt = 0 Then MsgBox DL_INFO1 & vbCrLf & TempPath & vbCrLf & vbCrLf & DL_INFO2, vbInformation, "Download Files"
            Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
            wsh.Run "msedge.exe " & url
            cnt = cnt + 1
        End If
    Next url
End Sub
Public Sub ExportModelCode(ByVal tPath As String)
    If OnnxModel Is Nothing Then Exit Sub
    WriteTextFile tPath, OnnxModel.JsCode
End Sub
Public Function Export(Optional ByRef target As Worksheet = Nothing, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
    If OnnxResults Is Nothing Then Exit Function
    If target Is Nothing Then Set target = ActiveSheet
    Set Export = OnnxModel.Export(target, Me, Left, Top)
End Function
Public Sub Render(ByRef target As GLFrame, Zoom As Double, offsetX As Double, offsetY As Double, pitch As Double, roll As Double, yaw As Double)
    Dim tH As Double, tW As Double, AspR As Double, Dst As Double, i As Long, yy As Double, bx As Double, by As Double, chr() As Byte
    Dim AoV As Double: AoV = 3
    With target
        tH = .height
        tW = .width
        AspR = tW / (tH + 0.0001)
        Dst = (0.5 * tH * Zoom) / Tan(AoV * 0.5 * DEG2RAD)
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
            .Viewport 0, 0, tW, tH
            .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
            .MatrixMode GL_PROJECTION
                .LoadIdentity
                .Perspective AoV, AspR, Dst * 0.5, Dst * 1.5
                .LookAt -offsetX, -offsetY - 0.001, Dst, -offsetX, -offsetY, 0, 0, 0, 1
            .MatrixMode GL_MODELVIEW
                .BindTexture GL_TEXTURE_2D, 1
                .LoadIdentity
                .PushMatrix
                    .Rotated pitch, -1, 0, 0
                    .Rotated roll, 0, -1, 0
                    .Rotated yaw, 0, 0, 1
                    .Color4f 1#, 1#, 1#, 1#
                    If Not OnnxResults Is Nothing Then
                        With OnnxResults
                            If .Count > 0 Then
                                If .Item(1).exists("error") = False Then OnnxModel.Render target, Me
                            End If
                        End With
                    End If
                .PopMatrix
            .MatrixMode GL_PROJECTION
                .LoadIdentity
                .Ortho2D 0, tW, 0, tH
                .MatrixMode GL_MODELVIEW
                .LoadIdentity
                bx = tW - BW - 30
                by = 10
                .Color4f 0, 1, 1, 1
                .Begin GL_LINE_LOOP
                    .Vertex2d bx, by
                    .Vertex2d bx + BW, by
                    .Vertex2d bx + BW, by + BH
                    .Vertex2d bx, by + BH
                .End1
                yy = by + BH - 20
                .listbase FONT_BASE_SMALL
                For i = LBound(glMsg) To UBound(glMsg)
                    chr = StrConv(CStr(glMsg(i)), vbFromUnicode)
                    .RasterPos2d bx + 10, yy
                    .CallLists UBound(chr) + 1, GL_UNSIGNED_BYTE, VarPtr(chr(0))
                    yy = yy - 18
                Next i
                .listbase 0
            .SwapBuffers
        End With
    End With
End Sub


