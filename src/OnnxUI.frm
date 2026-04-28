VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OnnxUI 
   Caption         =   "OnnxUI"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16905
   OleObjectBlob   =   "OnnxUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
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
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Const SEL_CL = &HFFCCCC
Private Const BASE_CL = &HFFFFFF
Private Const JS_EXPORT_MSG = "Do you want to export the model JavaScript code as a file?"
Private Const GWL_STYLE As Long = -16, WS_THICKFRAME = &H40000, ICO_SZ As Long = 36
Private WithEvents GLF As GLFrame
Attribute GLF.VB_VarHelpID = -1
Private WithEvents OnnxMain As VbOnnxMain
Attribute OnnxMain.VB_VarHelpID = -1
Private WithEvents rch As RichEdit
Attribute rch.VB_VarHelpID = -1
Private mform As Object
Private models As Object, busy As Boolean, slX As Double, slY As Double, pitch As Double, roll As Double, yaw As Double, zm As Double
Private Sub Frame2_Click(): DoEvents: End Sub
Private Sub rch_Change()
    OnnxMain.OnnxModel.editor.Value = rch.RichText
End Sub
Private Sub rch_Layout()
On Error GoTo err:
    With rch
        .height = mform.height - 120
        .width = mform.width - 25
    End With
err:
End Sub
Public Sub ShowModelForm()
On Error GoTo err:
    With OnnxMain
        Set mform = .GetFormFromName(TabStrip1.SelectedItem.Caption)
        Call Resizable(mform)
        Dim tbox As MSForms.TextBox: Set tbox = COnx(mform).editor
        tbox.Value = .OnnxModel.editor.Value
        Call ApplyRichEdit(tbox)
        mform.Show vbModal
    End With
err:
End Sub
Private Sub ApplyRichEdit(ByRef target As MSForms.TextBox)
    If rch Is Nothing Then Set rch = New RichEdit
    rch.Init target
End Sub
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
Private Sub TabStrip1_Change()
    If busy = True Then Exit Sub
    With OnnxMain
        Set .OnnxModel = .GetModelFromName(TabStrip1.SelectedItem.Caption)
        Set .OnnxResults = New Collection
        LabelInfo.Caption = .OnnxModel.INFO
        LabelModel.Caption = .OnnxModel.name
        LabelLibs.Caption = ""
        DoEvents
        .EnsureRuntimeFiles
        Call ResetCamera
    End With
End Sub
Private Sub LabelModel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If MsgBox(JS_EXPORT_MSG, vbYesNo) = vbYes Then
        Dim fpath As String: fpath = Application.GetSaveAsFilename(OnnxMain.OnnxModel.name & ".js", "javascript(*.js),*.js")
        If fpath Like "*.js" Then OnnxMain.ExportModelCode fpath
    End If
End Sub
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Label1
        busy = True
        .BackColor = SEL_CL
        Repaint
        Dim fpath As String: fpath = Application.GetOpenFilename(OnnxMain.FILE_FILTER)
        busy = False
        If fpath Like "*.*" Then ButtonAction Label1, "Execute", OnnxMain, fpath
        .BackColor = BASE_CL
    End With
End Sub
Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonAction Label2, "OpenTempFolder", OnnxMain
End Sub
Private Sub Label3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonAction Label3, "OpenEdgeForDownload", OnnxMain
End Sub
Private Sub Label4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonAction Label4, "Export", OnnxMain
End Sub
Private Sub Label5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LabelLibs.Caption = ""
    ButtonAction Label5, "EnsureRuntimeFiles", OnnxMain
End Sub
Private Sub Label6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
Private Sub ResetCamera()
    zm = 1: slX = 0: slY = 0: roll = 0: pitch = 0: yaw = 0
End Sub
Private Sub ButtonAction(ByRef target As Object, Optional FuncNm As String = "", Optional ByRef FuncObj As Object = Nothing, Optional ByRef args = Empty)
    If busy = True Then Exit Sub
    MousePointer = fmMousePointerHourGlass
    busy = True
    With target
        .BackColor = SEL_CL
        Repaint
        Sleep 150
        On Error GoTo err
        If FuncNm <> "" Then If IsEmpty(args) Then CallByName FuncObj, FuncNm, VbMethod Else CallByName FuncObj, FuncNm, VbMethod, args
        Sleep 150
err:
        MousePointer = fmMousePointerDefault
        .BackColor = BASE_CL
    End With
    busy = False
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
        Call Resizable(Me)
        DrawBuffer = 320000
        Dim key As Variant
        Sleep 500
        With TabStrip1
            For Each key In OnnxMain.ModelDict.keys()
                .Tabs.Add key
            Next key
            .BackColor = BASE_CL
        End With
        With Application.CommandBars
            Label1.Picture = .GetImageMso("PlayVideo", ICO_SZ, ICO_SZ)
            Label2.Picture = .GetImageMso("OpenFolder", ICO_SZ, ICO_SZ)
            Label3.Picture = .GetImageMso("WorkOffline", ICO_SZ, ICO_SZ)
            Label4.Picture = .GetImageMso("ExportExcel", ICO_SZ, ICO_SZ)
            Label5.Picture = .GetImageMso("RefreshAll", ICO_SZ, ICO_SZ)
            Label6.Picture = .GetImageMso("Info", ICO_SZ, ICO_SZ)
        End With
        Sleep 500
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
    Sleep 1000
End Sub
Private Sub UserForm_Terminate()
    Set GLF = Nothing
    Set OnnxMain = Nothing
    Set rch = Nothing
End Sub
Private Sub Resizable(ByRef fm As Variant)
    Dim hwnd As LongPtr, style As LongPtr
    WindowFromAccessibleObject fm, hwnd
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
