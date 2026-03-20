VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OnnxUI 
   Caption         =   "OnnxUI"
   ClientHeight    =   12075
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
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hWnd As LongPtr) As LongPtr
Private Const GWL_STYLE As Long = -16, WS_THICKFRAME = &H40000
Private hWnd As LongPtr, style As LongPtr
Private WithEvents OnnxMain As VbOnnxMain
Attribute OnnxMain.VB_VarHelpID = -1
Private WithEvents GLF As GLFrame
Attribute GLF.VB_VarHelpID = -1
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Const ICON_SIZE As Long = 36
Private models As Object
Private busy As Boolean
Private slideX As Double, slideY As Double, pitch As Double, roll As Double, yaw As Double, zm As Double

Private Sub UserForm_Resize()
    If busy Then Exit Sub
    On Error GoTo err
        Frame1.width = Me.width
        Frame1.height = Me.height - 261
        Frame2.Top = Me.height - 163
        Frame2.width = Me.width - 15.25
        Frame3.Left = Me.width - 263.25
        TabStrip1.width = Me.width - 11.25
err:
    Me.Repaint
    Call GLF_Paint
End Sub
Private Sub GLF_DblClick()
    Call ResetCamera
    Call GLF_Paint
End Sub
Private Sub GLF_DragDelta(ByVal DeltaX As Double, DeltaY As Double, Button As Integer)
    If busy Then Exit Sub
    If Button = 1 Then
        slideX = slideX - DeltaX
        slideY = slideY + DeltaY
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
Private Sub GLF_Paint()
    If busy Then Exit Sub
    Call OnnxMain.Render(GLF, zm, slideX, slideY, pitch, roll, yaw)
End Sub
Private Sub ButtonAction(ByRef target As Object, Optional FuncName As String = "", Optional ByRef FuncObject As Object = Nothing, Optional ByRef args = Empty)
    If busy = True Then Exit Sub
    Me.MousePointer = fmMousePointerHourGlass
    busy = True
    target.BackColor = &HFFCCCC
    Me.Repaint
    Sleep 150
    On Error GoTo err
    If FuncName <> "" Then
        If IsEmpty(args) Then CallByName FuncObject, FuncName, VbMethod Else CallByName FuncObject, FuncName, VbMethod, args
    End If
    Sleep 150
err:
    Me.MousePointer = fmMousePointerDefault
    target.BackColor = &HFFFFFF
    busy = False
    Me.Repaint
    Call GLF_Paint
End Sub
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    busy = True
    Label1.BackColor = &HFFCCCC
    Me.Repaint
    Dim fpath As String: fpath = Application.GetOpenFilename(OnnxMain.FILE_FILTER)
    busy = False
    If fpath Like "*.*" Then ButtonAction Me.Label1, "Execute", OnnxMain, fpath
    Label1.BackColor = &HFFFFFF
End Sub
Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Me.Label2, "OpenTempFolder", OnnxMain
End Sub
Private Sub Label3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Me.Label3, "OpenEdgeForDownload", OnnxMain
End Sub
Private Sub Label4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Me.Label4, "Export", OnnxMain
End Sub
Private Sub Label5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    LabelLibs.Caption = ""
    ButtonAction Me.Label5, "EnsureRuntimeFiles", OnnxMain
End Sub
Private Sub Label6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    ButtonAction Me.Label6, "ShowModelForm", Me
End Sub
Public Sub ShowModelForm()
    Dim tname As String: tname = TypeName(OnnxMain.OnnxModel)
    Dim obj As Object: Set obj = UserForms.Add(tname)
    obj.Show vbModal
End Sub
Private Sub LabelModel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If MsgBox("Do you want to export the model’s JavaScript code as a file?", vbYesNo) = vbYes Then
        Dim fpath As String
        fpath = Application.GetSaveAsFilename(OnnxMain.OnnxModel.name & ".js", "javascript(*.js),*.js")
        If fpath Like "*.js" Then
            OnnxMain.ExportModelCode fpath
        End If
    End If
End Sub
Private Sub UserForm_Initialize()
    Dim key As Variant
    busy = True
    Set OnnxMain = New VbOnnxMain
    Set models = OnnxMain.GetModelDict()
    Sleep 1000
    For Each key In models.keys()
        Me.TabStrip1.Tabs.Add key
    Next key
    Me.TabStrip1.BackColor = &HFFFFFF
    With Application.CommandBars
        Me.Label1.Picture = .GetImageMso("PlayVideo", ICON_SIZE, ICON_SIZE)
        Me.Label2.Picture = .GetImageMso("OpenFolder", ICON_SIZE, ICON_SIZE)
        Me.Label3.Picture = .GetImageMso("WorkOffline", ICON_SIZE, ICON_SIZE)
        Me.Label4.Picture = .GetImageMso("ExportExcel", ICON_SIZE, ICON_SIZE)
        Me.Label5.Picture = .GetImageMso("RefreshAll", ICON_SIZE, ICON_SIZE)
        Me.Label6.Picture = .GetImageMso("Info", ICON_SIZE, ICON_SIZE)
    End With
    busy = False
    Sleep 1000
    Call TabStrip1_Change
End Sub
Private Sub UserForm_Activate()
    If GLF Is Nothing Then
        Set GLF = New GLFrame
        GLF.Init Me.Frame1
        Call Resizable
    End If
End Sub
Private Sub OnnxMain_FileChecked(CheckedFilePath As String, url As String, Check As Boolean)
    With LabelLibs
        Dim chk As String: If Check Then chk = "[OK] : " Else chk = "[NG] : "
        .Caption = .Caption & chk & CheckedFilePath & " : " & url & vbNewLine
    End With
    Label1.Enabled = OnnxMain.ExecuteFlg
End Sub
Private Sub TabStrip1_Change()
    If busy = True Then Exit Sub
    With OnnxMain
        Set .OnnxModel = models.Item(TabStrip1.SelectedItem.Caption)
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
    zm = 1: slideX = 0: slideY = 0
    roll = 0: pitch = 0: yaw = 0
End Sub
Private Sub UserForm_Terminate()
    Set GLF = Nothing
    Set OnnxMain = Nothing
End Sub
Private Sub Resizable()
    WindowFromAccessibleObject Me, hWnd
    #If Win64 Then
        style = GetWindowLongPtr(hWnd, GWL_STYLE)
    #Else
        style = GetWindowLong(hWnd, GWL_STYLE)
    #End If
    style = (style Or WS_THICKFRAME Or &H30000) 'And Not WS_CAPTION
    #If Win64 Then
        SetWindowLongPtr hWnd, GWL_STYLE, style
    #Else
        SetWindowLong hWnd, GWL_STYLE, style
    #End If
End Sub
