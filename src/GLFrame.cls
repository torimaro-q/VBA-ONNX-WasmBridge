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
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Const Inv255 As Double = 1 / 255
Public Event Load()
Public Event Paint()
Public Event DblClick()
Public Event Click(ByVal X As Double, y As Double, Button As Integer)
Public Event DragDelta(ByVal DeltaX As Double, DeltaY As Double, Button As Integer)
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAcessible As Object, ByRef hWnd As LongPtr) As LongPtr
Private px As Double, pY As Double, Tw2Px As Double
Public gl As OpenGL
Private Sub Class_Terminate()
    gl.PaintEnd
    Set gl = Nothing
End Sub
Private Sub Frame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClick
End Sub
Private Sub Frame_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal y As Single)
    px = Tw2Px * X
    pY = Tw2Px * y
    RaiseEvent Click(px, pY, Button)
End Sub
Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button = 0 Then Exit Sub
    RaiseEvent DragDelta(px - Tw2Px * X, pY - Tw2Px * y, Button)
    px = Tw2Px * X: pY = Tw2Px * y
End Sub
Public Sub Refresh()
    RaiseEvent Paint
End Sub
Public Property Get width() As Double
    If Frame Is Nothing Then Exit Sub
    width = Frame.width * Tw2Px
End Property
Public Property Get height() As Double
    If Frame Is Nothing Then Exit Sub
    height = Frame.height * Tw2Px
End Property
Public Sub Init(ByRef TargetFrame As MSForms.Frame)
    Dim hw As LongPtr
    Set gl = New OpenGL
    Tw2Px = GetDPI() / 72
    DoEvents
    WindowFromAccessibleObject TargetFrame, hw
    With TargetFrame
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Dim bColor As Color4
    bColor = Rgb2Color4(TargetFrame.BackColor)
    With gl
        .hWnd = hw
        .PaintStart
        .ClearColor bColor.R, bColor.G, bColor.b, bColor.a
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
        .b = rgbcl.b * Inv255
        .a = 1
    End With
End Function
