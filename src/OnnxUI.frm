VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OnnxUI 
   Caption         =   "OnnxUI"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16905
   OleObjectBlob   =   "OnnxUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'ÉIÅ[ÉiÅ[ ÉtÉHÅ[ÉÄÇÃíÜâõ
End
Attribute VB_Name = "OnnxUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hwnd As LongPtr) As LongPtr
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Const GWL_STYLE As Long = -16, WS_THICKFRAME = &H40000, ICO_SZ As Long = 36
Private hwnd As LongPtr, style As LongPtr
Private WithEvents OnnxMain As VbOnnxMain
Attribute OnnxMain.VB_VarHelpID = -1
Private WithEvents GLF As GLFrame
Attribute GLF.VB_VarHelpID = -1
Private models As Object, busy As Boolean, slX As Double, slY As Double, pitch As Double, roll As Double, yaw As Double, zm As Double
Private rch As RichEdit
Private Sub ApplyRichEdit(ByRef target As MSForms.TextBox)
    If rch Is Nothing Then Set rch = New RichEdit
    rch.Init target
End Sub
Private Sub Frame2_Click(): DoEvents: End Sub
Private Sub GLF_Paint()
    MousePointer = fmMousePointerHourGlass
    Call OnnxMain.Render(GLF, zm, slX, slY, pitch, roll, yaw)
    MousePointer = fmMousePointerDefault
End Sub
Private Sub OnnxMain_FileChecked(CheckedFilePath As String, url As String, Check As Boolean)
    With LabelLibs
        Dim chk As String: If Check Then chk = "[OK] : " Else chk = "[NG] : "
        .Caption = .Caption & chk & CheckedFilePath & " : " & url & vbNewLine
    End With
    Label1.Enabled = OnnxMain.ExecuteFlg
End Sub
Private Sub GLF_DragDelta(ByVal DeltaX As Double, DeltaY As Double, Button As Integer)
    If busy Then Exit Sub
    If Button = 1 Then
        slX = slX - DeltaX
        slY = slY + DeltaY
    ElseIf Button = 2 Then
        zm = zm + DeltaY * 0.001
    ElseIf Button = 3 Then
        roll = roll + DeltaX
        pitch = pitch + DeltaY
    Else
        yaw = yaw + DeltaY
    End If
    Call GLF_Paint
End Sub
Public Sub ShowModelForm()
    Dim tname As String: tname = TypeName(OnnxMain.OnnxModel)
    Dim obj As Object: Set obj = UserForms.Add(tname)
    ApplyRichEdit OnnxMain.COnx(obj).Editor
    obj.Show vbModal
End Sub
Private Sub LabelModel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If MsgBox("Do you want to export the modelÔøΩfs JavaScript code as a file?", vbYesNo) = vbYes Then
        Dim fpath As String: fpath = Application.GetSaveAsFilename(OnnxMain.OnnxModel.name & ".js", "javascript(*.js),*.js")
        If fpath Like "*.js" Then OnnxMain.ExportModelCode fpath
    End If
End Sub
Private Sub TabStrip1_Change()
    If busy = True Then Exit Sub
    With OnnxMain
        Set .OnnxModel = .COnx(models.Item(CStr(TabStrip1.SelectedItem.Caption)))
        Set OnnxMain.OnnxResults = New Collection
        LabelInfo.Caption = .OnnxModel.Info
        LabelModel.Caption = .OnnxModel.name
        LabelLibs.Caption = ""
        DoEvents
        .EnsureRuntimeFiles
        Call ResetCamera
    End With
End Sub
Private Sub ResetCamera()
    zm = 1: slX = 0: slY = 0: roll = 0: pitch = 0: yaw = 0
End Sub
Private Sub ButtonAction(ByRef target As Object, Optional FuncName As String = "", Optional ByRef FuncObject As Object = Nothing, Optional ByRef args = Empty)
    If busy = True Then Exit Sub
    MousePointer = fmMousePointerHourGlass
    busy = True
    target.BackColor = &HFFCCCC
    Repaint
    Sleep 150
    On Error GoTo err
    If FuncName <> "" Then
        If IsEmpty(args) Then CallByName FuncObject, FuncName, VbMethod Else CallByName FuncObject, FuncName, VbMethod, args
    End If
    Sleep 150
err:
    MousePointer = fmMousePointerDefault
    target.BackColor = &HFFFFFF
    busy = False
    Repaint
    Call GLF_Paint
End Sub
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    With Label1
        busy = True
        .BackColor = &HFFCCCC
        Repaint
        Dim fpath As String: fpath = Application.GetOpenFilename(OnnxMain.FILE_FILTER)
        busy = False
        If fpath Like "*.*" Then ButtonAction Label1, "Execute", OnnxMain, fpath
        .BackColor = &HFFFFFF
    End With
End Sub
Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Label2, "OpenTempFolder", OnnxMain
End Sub
Private Sub Label3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Label3, "OpenEdgeForDownload", OnnxMain
End Sub
Private Sub Label4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Label4, "Export", OnnxMain
End Sub
Private Sub Label5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    LabelLibs.Caption = ""
    ButtonAction Label5, "EnsureRuntimeFiles", OnnxMain
End Sub
Private Sub Label6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Label6, "ShowModelForm", Me
End Sub
Private Sub UserForm_Resize()
    If busy Then Exit Sub
    On Error GoTo err
        With Frame1
            .width = width
            .height = height - 281
        End With
        With Frame2
            .Top = height - 183
            .width = width - 15.25
        End With
        Frame3.Left = width - 263.25
        TabStrip1.width = width - 11.25
err:
    Repaint
    Call GLF_Paint
End Sub
Private Sub GLF_DblClick()
    If busy Then Exit Sub
    Call ResetCamera
    Call GLF_Paint
End Sub
Private Sub UserForm_Activate()
    busy = True
    If GLF Is Nothing Then
        Call Resizable
        DrawBuffer = 320000
        Sleep 1000
        Dim key As Variant
        For Each key In models.keys()
            TabStrip1.Tabs.Add key
        Next key
        TabStrip1.BackColor = &HFFFFFF
        With Application.CommandBars
            Label1.Picture = .GetImageMso("PlayVideo", ICO_SZ, ICO_SZ)
            Label2.Picture = .GetImageMso("OpenFolder", ICO_SZ, ICO_SZ)
            Label3.Picture = .GetImageMso("WorkOffline", ICO_SZ, ICO_SZ)
            Label4.Picture = .GetImageMso("ExportExcel", ICO_SZ, ICO_SZ)
            Label5.Picture = .GetImageMso("RefreshAll", ICO_SZ, ICO_SZ)
            Label6.Picture = .GetImageMso("Info", ICO_SZ, ICO_SZ)
        End With
        DoEvents
        Set GLF = New GLFrame
        GLF.Init Frame1
    End If
    busy = False
    Call TabStrip1_Change
    DoEvents
End Sub
Private Sub UserForm_Initialize()
    Set OnnxMain = New VbOnnxMain
    Set models = OnnxMain.GetModelDict()
End Sub
Private Sub UserForm_Terminate()
    Set GLF = Nothing
    Set OnnxMain = Nothing
End Sub
Private Sub Resizable()
    WindowFromAccessibleObject Me, hwnd
    #If Win64 Then
        style = GetWindowLongPtr(hwnd, GWL_STYLE)
    #Else
        style = GetWindowLong(hwnd, GWL_STYLE)
    #End If
    style = (style Or WS_THICKFRAME Or &H30000)
    #If Win64 Then
        SetWindowLongPtr hwnd, GWL_STYLE, style
    #Else
        SetWindowLong hwnd, GWL_STYLE, style
    #End If
End Sub
