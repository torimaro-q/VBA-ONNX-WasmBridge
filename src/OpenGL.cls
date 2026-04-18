VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal format As Long, ByRef ppfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SwapBuffersA Lib "gdi32" Alias "SwapBuffers" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub gluPerspective Lib "glu32.dll" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
Private Declare PtrSafe Sub gluLookAt Lib "glu32.dll" (ByVal eyex As Double, ByVal eyey As Double, ByVal eyez As Double, ByVal centerx As Double, ByVal centery As Double, ByVal centerz As Double, ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
Private Declare PtrSafe Sub gluOrtho2D Lib "glu32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double)
Private Declare PtrSafe Sub glOrtho Lib "opengl32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zbear As Double, ByVal zFar As Double)
Private Declare PtrSafe Sub glBlendFunc Lib "opengl32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long)
Private Declare PtrSafe Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal hGLRC As LongPtr) As Long
Private Declare PtrSafe Function wglDeleteContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Sub glColor4f Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColor3d Lib "opengl32.dll" (ByVal red As Double, ByVal green As Double, ByVal blue As Double)
Private Declare PtrSafe Sub glVertex2d Lib "opengl32.dll" (ByVal a As Double, ByVal b As Double)
Private Declare PtrSafe Sub glVertex3d Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glNormal3d Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glVertex3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glNormal3dv Lib "opengl32.dll" (ByVal dPtr As LongPtr)
Private Declare PtrSafe Sub glBegin Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glEnd Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glMultMatrixd Lib "opengl32.dll" (ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal X As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
Private Declare PtrSafe Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
Private Declare PtrSafe Sub glEnableClientState Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glDisableClientState Lib "opengl32.dll" (ByVal GLenum As Long)
Private Declare PtrSafe Sub glMaterialfv Lib "opengl32.dll" (ByVal Face As Long, ByVal pname As Long, ByVal param As LongPtr)
Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal Mask As Long)
Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
Private Declare PtrSafe Sub glColorPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glNormalPointer Lib "opengl32.dll" (ByVal GLenum As Long, ByVal Stride As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glVertexPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glDrawArrays Lib "opengl32.dll" (ByVal GLenum As Long, ByVal first As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDrawElements Lib "opengl32.dll" (ByVal GLenum As Long, ByVal Count As Long, ByVal GlenumT As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glClipPlane Lib "opengl32.dll" (ByVal GLenum As Long, ByVal ptr As LongPtr)

Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr

Private Declare PtrSafe Function wglUseFontBitmaps Lib "opengl32.dll" Alias "wglUseFontBitmapsA" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long
Private Declare PtrSafe Function wglUseFontBitmapsW Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal dfirst As Long, ByVal dcount As Long, ByVal dlistbase As Long) As Long

Private Declare PtrSafe Sub glRasterPos3d Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glRasterPos2d Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double)
Private Declare PtrSafe Sub glTranslated Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glRotated Lib "opengl32.dll" (ByVal angle As Double, ByVal X As Double, ByVal y As Double, ByVal z As Double)
Private Declare PtrSafe Sub glScaled Lib "opengl32.dll" (ByVal X As Double, ByVal y As Double, ByVal z As Double)

Private Declare PtrSafe Function glIsList Lib "opengl32.dll" (ByVal index As Long) As Byte
Private Declare PtrSafe Sub glNewList Lib "opengl32.dll" (ByVal index As Long, ByVal GLenum As Long)
Private Declare PtrSafe Sub glEndList Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glCallList Lib "opengl32.dll" (ByVal index As Long)
Private Declare PtrSafe Sub glDeleteLists Lib "opengl32.dll" (ByVal index As Long, ByVal size As Long)
Private Declare PtrSafe Sub glDeleteList Lib "opengl32.dll" (ByVal index As Long)
Private Declare PtrSafe Sub glListBase Lib "opengl32.dll" (ByVal base As Long)
Private Declare PtrSafe Sub glCallLists Lib "opengl32.dll" (ByVal size As Long, ByVal vType As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glPixelStorei Lib "opengl32.dll" (ByVal GLenum As Long, ByVal param As Long)
Private Declare PtrSafe Sub glTexParameteri Lib "opengl32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long, ByVal param As Long)
Private Declare PtrSafe Sub glTexCoordPointer Lib "opengl32.dll" (ByVal size As Long, ByVal GLenum As Long, ByVal Stride As Long, ByVal ptr As LongPtr)
Private Declare PtrSafe Sub glTexImage2D Lib "opengl32.dll" (ByVal GLenum As Long, ByVal level As Long, ByVal internalformat As Long, _
                                                             ByVal width As Long, ByVal height As Long, ByVal border As Long, _
                                                             ByVal GlenumF As Long, ByVal GlenumT As Long, ByVal PixelPtr As LongPtr)
Private Declare PtrSafe Sub gluBuild2DMipmaps Lib "glu32.dll" (ByVal GLenum1 As Long, ByVal GLenum2 As Long, _
                                                              ByVal width As Long, ByVal height As Long, _
                                                              ByVal GLenum3 As Long, ByVal GLenum4 As Long, ByVal PixelPtr As LongPtr)
Private Declare PtrSafe Sub glTexCoord2d Lib "opengl32.dll" (ByVal s As Double, ByVal t As Double)

Private Declare PtrSafe Sub glGenTextures Lib "opengl32.dll" (ByVal n As Long, ByRef textures As Long)
Private Declare PtrSafe Sub glBindTexture Lib "opengl32.dll" (ByVal target As Long, ByVal texture As Long)
Private Declare PtrSafe Sub glDeleteTextures Lib "opengl32.dll" (ByVal n As Long, ByRef textures As Long)

Private pfd As PIXELFORMATDESCRIPTOR
Public hformat As Long, hGLRC As LongPtr, hdc As LongPtr
Private myhwnd As LongPtr
Public Sub CreateFontList(Optional ByVal FontSize As Long = 30, Optional ByVal FontName As String = "Arial", Optional ByVal weight As Long = 200, Optional listbase As Long = 1000, Optional genCount As Long = 255, Optional isUnicode As Boolean = False)
    Dim hFont As LongPtr, hOldFont As LongPtr
    If glIsList(listbase) = 1 Then glDeleteLists listbase, genCount
    If isUnicode = False Then
        hFont = CreateFont(FontSize, 0, 0, 0, weight, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, FontName)
    Else
        hFont = CreateFontW(FontSize, 0, 0, 0, weight, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, FontName)
    End If
    If hFont = 0 Then Exit Sub
    hOldFont = SelectObject(hdc, hFont)
    If isUnicode = False Then
        wglUseFontBitmaps hdc, 0, genCount, listbase
    Else
        wglUseFontBitmapsW hdc, 0, genCount, listbase
    End If
    Call SelectObject(hdc, hOldFont)
    Call DeleteObject(hFont)
End Sub
Public Sub GenTextures(ByVal n As Long, ByRef textures As Long)
    glGenTextures n, textures
End Sub
Public Sub BindTexture(ByVal target As GLenum, ByRef textures As Long)
    glBindTexture target, textures
End Sub
Public Sub DeleteTextures(ByVal n As Long, ByRef textures As Long)
    glDeleteTextures n, textures
End Sub
Public Sub DisposeFontList()
    If glIsList(FONT_BASE_SMALL) = 1 Then glDeleteLists FONT_BASE_SMALL, FONT_BASE_UNICODE + 65535
End Sub
Public Sub NewList(ByVal index As Long, CompileMode As GLenum)
    Call glNewList(index, CompileMode)
End Sub
Public Sub EndList()
    Call glEndList
End Sub
Public Sub CallList(ByVal index As Long)
    Call glCallList(index)
End Sub
Public Sub CallLists(ByVal index As Long, ByVal vType As GLenum, ByVal ptr As LongPtr)
    Call glCallLists(index, vType, ptr)
End Sub
Public Sub DeleteLists(ByVal index As Long, size As Long)
    Call glDeleteLists(index, size)
End Sub
Public Sub listbase(ByVal base As Long)
    Call glListBase(base)
End Sub
Public Function IsList(ByVal index As Long) As Boolean
    IsList = glIsList(index)
End Function
Public Sub RasterPos3d(ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glRasterPos3d(X, y, z)
End Sub
Public Sub RasterPos2d(ByVal X As Double, ByVal y As Double)
    Call glRasterPos2d(X, y)
End Sub
Public Sub Scaled(ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glScaled(X, y, z)
End Sub
Public Sub Translated(ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glTranslated(X, y, z)
End Sub
Public Sub PointSize(ByVal size As Single)
    Call glPointSize(size)
End Sub
Public Sub LineWidth(ByVal size As Single)
    Call glLineWidth(size)
End Sub
Public Sub EnableClientState(ByVal Gle As GLenum)
    Call glEnableClientState(Gle)
End Sub
Public Sub DisableClientState(ByVal Gle As GLenum)
    Call glDisableClientState(Gle)
End Sub
Public Sub ColorPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptr As LongPtr)
    Call glColorPointer(size, Gle, Stride, ptr)
End Sub
Public Sub VertexPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptr As LongPtr)
    Call glVertexPointer(size, Gle, Stride, ptr)
End Sub
Public Sub NormalPointer(ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptr As LongPtr)
    Call glNormalPointer(Gle, Stride, ptr)
End Sub
Public Sub DrawArrays(ByVal Mode As GLenum, ByVal first As Long, ByVal size As Long)
    Call glDrawArrays(Mode, first, size)
End Sub
Public Sub DrawElements(ByVal Mode As GLenum, ByVal Count As Long, ByVal Type1 As GLenum, ByVal ptr As LongPtr)
    Call glDrawElements(Mode, Count, Type1, ptr)
End Sub
Public Sub MultMatrixd(ByVal ptr As LongPtr)
    Call glMultMatrixd(ptr)
End Sub
Public Sub PaintStart()
    Call wglMakeCurrent(hdc, hGLRC)
    DoEvents
    Call CreateFontList(20, "Impact", 100, FONT_BASE_SMALL)
    Call CreateFontList(40, "Impact", 300, FONT_BASE_NORMAL)
    Call CreateFontList(60, "Impact", 300, FONT_BASE_LARGE)
    Call CreateFontList(80, "Impact", 600, FONT_BASE_EXLARGE)
End Sub
Public Property Get UseUnicode() As Boolean
    If glIsList(FONT_BASE_UNICODE) = 1 Then UseUnicode = True Else UseUnicode = False
End Property
Public Property Let UseUnicode(RHS As Boolean)
    If RHS Then
        If glIsList(FONT_BASE_UNICODE) <> 1 Then Call CreateFontList(24, "Meiryo UI", 1000, FONT_BASE_UNICODE, 65535, True)
    Else
        If glIsList(FONT_BASE_UNICODE) = 1 Then glDeleteLists FONT_BASE_UNICODE, 65535
    End If
End Property
Public Sub PaintEnd()
    Call SwapBuffersA(hdc)
    Call wglMakeCurrent(0, 0)
End Sub
Public Sub MakeCurrent(ByVal hdc As LongPtr, ByVal hGLRC As LongPtr)
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub SwapBuffers()
    Call SwapBuffersA(hdc)
End Sub
Public Sub Viewport(ByVal X As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Call glViewport(X, y, width, height)
End Sub
Public Sub Enable(ByVal Gle As GLenum)
    Call glEnable(Gle)
End Sub
Public Sub Disable(ByVal Gle As GLenum)
    Call glDisable(Gle)
End Sub
Public Sub MatrixMode(ByVal Gle As GLenum)
    Call glMatrixMode(Gle)
End Sub
Public Sub Materialfv(ByVal GLe1 As GLenum, ByVal GLe2 As GLenum, ByVal param As LongPtr)
    Call glMaterialfv(GLe1, GLe2, param)
End Sub
Public Sub LoadIdentity()
    Call glLoadIdentity
End Sub
Public Sub PushMatrix()
    Call glPushMatrix
End Sub
Public Sub PopMatrix()
    Call glPopMatrix
End Sub
Public Sub Begin(ByVal Gle As GLenum)
    Call glBegin(Gle)
End Sub
Public Sub End1()
    Call glEnd
End Sub
Public Sub Color4f(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Call glColor4f(red, green, blue, alpha)
End Sub
Public Sub Color3d(ByVal red As Double, ByVal green As Double, ByVal blue As Double)
    Call glColor3d(red, green, blue)
End Sub
Public Sub Vertex2d(ByVal a As Double, ByVal b As Double)
    Call glVertex2d(a, b)
End Sub
Public Sub Vertex3d(ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glVertex3d(X, y, z)
End Sub
Public Sub Normal3d(ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glNormal3d(X, y, z)
End Sub
Public Sub Vertex3dv(ByVal Vec4Ptr As LongPtr)
    Call glVertex3dv(Vec4Ptr)
End Sub
Public Sub Normal3dv(ByVal Vec3Ptr As LongPtr)
    Call glNormal3dv(Vec3Ptr)
End Sub
Public Sub Rotated(ByVal angle As Double, ByVal X As Double, ByVal y As Double, ByVal z As Double)
    Call glRotated(angle, X, y, z)
End Sub
Public Sub BlendFunc(ByVal GLe1 As GLenum, ByVal GLe2 As GLenum)
    Call glBlendFunc(GLe1, GLe2)
End Sub
Public Sub Clear(Optional ByVal Mask As Long = 16384)
    Call glClear(Mask)
End Sub
Public Sub ClipPlane(ByVal plane As GLenum, ByVal Vector4dPtr As LongPtr)
     Call glClipPlane(plane, Vector4dPtr)
End Sub
Public Sub ClearColor(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Call glClearColor(red, green, blue, alpha)
End Sub
Public Sub Perspective(ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double)
    Call gluPerspective(fovy, aspect, zNear, zFar)
End Sub
Public Sub Ortho2D(ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double)
    Call gluOrtho2D(Left, Right, Bottom, Top)
End Sub
Public Sub Ortho(ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Call glOrtho(Left, Right, Bottom, Top, zNear, zFar)
End Sub
Public Sub LookAt(ByVal eyex As Double, ByVal eyey As Double, ByVal eyez As Double, ByVal centerx As Double, ByVal centery As Double, ByVal centerz As Double, ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
    Call gluLookAt(eyex, eyey, eyez, centerx, centery, centerz, upx, upy, upz)
End Sub
Public Sub PixelStorei(ByVal pname As GLenum, ByVal param As Long)
    Call glPixelStorei(pname, param)
End Sub
Public Sub TexParameteri(ByVal target As GLenum, ByVal pname As GLenum, ByVal param As Long)
    Call glTexParameteri(target, pname, param)
End Sub
Public Sub TexImage2D(ByVal target As GLenum, ByVal level As Long, ByVal internalformat As Long, ByVal width As Long, ByVal height As Long, ByVal border As Long, ByVal format As GLenum, ByVal pType As GLenum, ByVal PixelPtr As LongPtr)
    Call glTexImage2D(target, level, internalformat, width, height, border, format, pType, PixelPtr)
End Sub
Public Sub Build2DMipmaps(ByVal target As GLenum, ByVal internalformat As GLenum, ByVal width As Long, ByVal height As Long, ByVal format As GLenum, ByVal pType As GLenum, ByVal PixelPtr As LongPtr)
    Call gluBuild2DMipmaps(target, internalformat, width, height, format, pType, PixelPtr)
End Sub
Public Sub TexCoordPointer(ByVal size As Long, ByVal Gle As GLenum, ByVal Stride As Long, ByVal ptr As LongPtr)
    Call glTexCoordPointer(size, Gle, Stride, ptr)
End Sub
Public Sub TexCoord2d(ByVal s As Double, ByVal t As Double)
    Call glTexCoord2d(s, t)
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let hwnd(newPtr As LongPtr)
    myhwnd = newPtr
    Call Inithwnd
End Property
Public Property Get hwnd() As LongPtr
    hwnd = myhwnd
End Property
Private Sub Inithwnd()
    If hdc <> 0 Then Releasehwnd
    hdc = GetDC(myhwnd)
    hformat = ChoosePixelFormat(hdc, pfd)
    Call SetPixelFormat(hdc, hformat, pfd)
    hGLRC = wglCreateContext(hdc)
End Sub
Private Sub Releasehwnd()
    Call DisposeFontList
    Call wglMakeCurrent(0, 0)
    Call wglDeleteContext(hGLRC)
    Call ReleaseDC(myhwnd, hdc)
End Sub
Private Sub Class_Initialize()
    With pfd
        .nSize = 40 ' LenB(pfd)
        .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_SUPPORT_GDI 'GDI
        .cColorBits = 32
        .cRedBits = 0
        .cRedShift = 0
        .cGreenBits = 0
        .cGreenShift = 0
        .cBlueBits = 0
        .cBlueShift = 0
        .cAlphaBits = 0
        .cAlphaShift = 0
        .cAccumBits = 0
        .cAccumRedBits = 0
        .cAccumGreenBits = 0
        .cAccumBlueBits = 0
        .cAccumAlphaBits = 0
        .cDepthBits = 32
        .cStencilBits = 8
        .cAuxBuffers = 0
        .iLayerType = 0 'PFD_MAIN_PLANE
        .bReserved = 0
        .dwLayerMask = 0
        .dwVisibleMask = 0
        .dwDamageMask = 0
    End With
End Sub
Private Sub Class_Terminate()
    Call Releasehwnd
End Sub


