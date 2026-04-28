VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Frame As MSForms.Frame
Attribute Frame.VB_VarHelpID = -1
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Public Event Load()
Public Event Paint()
Public Event DblClick()
Public Event Click(ByVal X As Double, Y As Double, Button As Integer)
Public Event DragDelta(ByVal DeltaX As Double, DeltaY As Double, Button As Integer)
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAcessible As Object, ByRef hwnd As LongPtr) As LongPtr
Private pX As Double, pY As Double, TW2PX As Double
Public gl As OpenGL
Private Sub Class_Terminate()
    gl.PaintEnd
    Set gl = Nothing
End Sub
Private Sub Frame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClick
End Sub
Private Sub Frame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pX = TW2PX * X
    pY = TW2PX * Y
    RaiseEvent Click(pX, pY, Button)
End Sub
Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 0 Then Exit Sub
    RaiseEvent DragDelta(pX - TW2PX * X, pY - TW2PX * Y, Button)
    pX = TW2PX * X: pY = TW2PX * Y
End Sub
Public Sub Refresh()
    RaiseEvent Paint
End Sub
Public Property Get width() As Double
    If Frame Is Nothing Then Exit Sub
    width = Frame.width * TW2PX
End Property
Public Property Get height() As Double
    If Frame Is Nothing Then Exit Sub
    height = Frame.height * TW2PX
End Property
Public Sub Init(ByRef TargetFrame As MSForms.Frame)
    Dim hw As LongPtr
    Set gl = New OpenGL
    TW2PX = GetDPI() / 72
    DoEvents
    WindowFromAccessibleObject TargetFrame, hw
    With TargetFrame
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Dim bColor As Color4
    bColor = Rgb2Color4(TargetFrame.BackColor)
    With gl
        .hwnd = hw
        .PaintStart
        .ClearColor bColor.R, bColor.G, bColor.B, bColor.A
        .Enable GL_DEPTH_TEST
    End With
    Set Frame = TargetFrame
    DoEvents
    RaiseEvent Load
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Public Function Rgb2Color4(ByVal sysCL As Long) As Color4
    Dim rgbcl As RGBA, raw As L, cl As Color4
    If (sysCL And CLMASK) <> 0 Then raw.v = GetSysColor(CLng(sysCL And &HFF&)) Else raw.v = sysCL
    LSet rgbcl = raw
    With Rgb2Color4
        .R = rgbcl.R * Inv255
        .G = rgbcl.G * Inv255
        .B = rgbcl.B * Inv255
        .A = 1
    End With
End Function
