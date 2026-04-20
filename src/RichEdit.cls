VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private myCtr As MSForms.Control
Private WithEvents myFrm As MSForms.Frame
Attribute myFrm.VB_VarHelpID = -1
Private WithEvents myTbx As MSForms.TextBox
Attribute myTbx.VB_VarHelpID = -1
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As LongPtr, pSource As LongPtr, ByVal dwLength As Long)
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long
Private Const SCF_SELECTION  As Long = &H1&, LOGPIXELSX As Long = 88, SCF_ALL = &H4
Private Const CFM_COLOR  As Long = &H40000000, CFM_WEIGHT  As Long = 4194304
Private Const CFM_CHARSET  As Long = 134217728, CFM_FACE As Long = 536870912
Private Const ES_MULTILINE As Long = &H4&, ES_AUTOVSCROLL As Long = &H40&, ES_AUTOHSCROLL  As Long = &H80&, ES_DISABLENOSCROLL = &H2000
Private Const RDW_INVALIDATE As Long = &H1, RDW_UPDATENOW As Long = &H100, RDW_ERASE As Long = &H4
Private Const WS_CHILD  As Long = &H40000000, WS_VISIBLE  As Long = &H10000000
Private Const WM_SETTEXT  As Long = &HC&, WM_GETTEXT  As Long = &HD&, WM_USER = &H400
Private Const WM_GETTEXTLENGTH As Long = &HE&, WM_SETREDRAW  As Long = &HB, WS_VSCROLL = &H200000
Private Const EM_SETCHARFORMAT As Long = &H444&, EM_GETSEL  As Long = &HB0&, EM_SETSEL  As Long = &HB1&
Private Const EM_SETSCROLLPOS = WM_USER + 222, EM_GETSCROLLPOS = WM_USER + 221, EM_SETBKGNDCOLOR = &H443
Private Const EM_REPLACESEL = &HC2, EM_GETLINECOUNT = &HBA, EM_LINEINDEX = &HBB, EM_LINELENGTH = &HC1
Private Const EM_GETLINE = &HC4, EM_LINEFROMCHAR = &HC9
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long '8
    dwEffects As Long '12
    yHeight As Long '16
    yOffset As Long '20
    crTextColor As Long '24
    bCharSet As Byte '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To 64 - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lcid As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bUnderlineColor As Byte
End Type
Private Const MY_RICH_EDIT_STYLE As Long = WS_CHILD Or WS_VISIBLE Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or WS_VSCROLL Or ES_DISABLENOSCROLL
Private Const MY_MASK As Long = CFM_FACE Or CFM_COLOR
Private MY_FONT_COLOR  As Long, MY_BACK_COLOR As Long, TW2PX As Double
Private Const PAD As Long = 2
Public PtnDic As Object
Private IsInited As Boolean, Reg As Object, LibPtr As LongPtr, CF2 As CHARFORMAT2
Private rHw As LongPtr, fHw As LongPtr, tHw As LongPtr, buf As String, pre As String
Private Sub SetColor(ByRef startPos As Long, ByRef length As Long, ByRef color As Long)
    CF2.crTextColor = color
    Dim rh As LongPtr: rh = RichHw
    Call SendMessage(rh, EM_SETSEL, startPos, startPos + length)
    Call SendMessage(rh, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(CF2))
End Sub
Private Sub myFrm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        Call Highlight
        buf = Me.RichText
        If buf <> pre Then
            myTbx.Value = buf
            pre = buf
        End If
        KeyCode = 0
    End If
End Sub
Private Sub myFrm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        If Shift = 0 Then SendMessage RichHw, EM_REPLACESEL, True, StrPtr(vbTab)
        KeyCode = 0
    End If
End Sub
Private Sub Highlight()
    Dim txt As String: txt = Me.RichText
    Dim m As Object, s As Long, e As Long, pkey As Variant, cl As Long, scrollPos As POINTAPI
    Dim rh As LongPtr: rh = RichHw
    SendMessage rh, WM_SETREDRAW, 0, 0
    SendMessage rh, EM_GETSEL, VarPtr(s), VarPtr(e)
    SendMessage rh, EM_GETSCROLLPOS, 0, VarPtr(scrollPos)
    SetColor 0, Len(txt), MY_FONT_COLOR
    With Reg
        For Each pkey In PtnDic.keys
            .Pattern = pkey
            cl = PtnDic.Item(pkey)
            For Each m In .Execute(txt)
                With m
                    SetColor .FirstIndex, .length, cl
                End With
            Next m
        Next pkey
    End With
    SendMessage rh, EM_SETSCROLLPOS, 0, VarPtr(scrollPos)
    SendMessage rh, EM_SETSEL, s, e
    SendMessage rh, WM_SETREDRAW, 1, 0
    RedrawWindow rh, 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW
End Sub
Public Property Get RichHw()
    If IsInited = False Then Exit Property
    With myCtr
        tHw = .[_GethWnd]
        If fHw <> tHw Then
            DestroyWindow rHw
            fHw = tHw
            rHw = SummonEditor
            Me.RichText = myTbx.Value
        End If
        RichHw = rHw
        myFrm.SetFocus
    End With
End Property
Public Property Get RichText() As Variant
    Dim rh As LongPtr: rh = RichHw
    If rh = 0 Then Exit Property
    Dim tLen As Long: tLen = CLng(SendMessage(rh, WM_GETTEXTLENGTH, 0, 0))
    If tLen > 0 Then
        Dim buffer As String: buffer = String$(tLen + 1, vbNullChar)
        Call SendMessage(rh, WM_GETTEXT, tLen + 1, StrPtr(buffer))
        RichText = Replace(Left$(buffer, tLen), vbNewLine, vbLf)
    Else
        RichText = ""
    End If
End Property
Public Property Let RichText(newText As Variant)
    Dim rh As LongPtr: rh = RichHw
    If rh = 0 Then Exit Property
    Call SendMessage(rh, WM_SETTEXT, 0, StrPtr(CStr(newText)))
End Property
Private Sub myFrm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If RichHw = 0 Then Exit Sub
End Sub
Private Sub myFrm_Layout()
    With myFrm
        Call MoveWindow(RichHw, PAD, PAD, CLng(TW2PX * .width - PAD * 2), CLng(TW2PX * .height - PAD * 2), 1)
    End With
End Sub
Private Function SummonEditor() As LongPtr
    With myCtr
        SummonEditor = CreateWindowEx(0, "RICHEDIT50W", .Tag, MY_RICH_EDIT_STYLE, PAD, PAD, CLng(TW2PX * .width - PAD * 2), CLng(TW2PX * .height - PAD * 2), fHw, 0, 0, 0)
        SendMessage SummonEditor, EM_SETCHARFORMAT, 0, VarPtr(CF2)
        SendMessage SummonEditor, EM_SETBKGNDCOLOR, 0, MY_BACK_COLOR
    End With
    Call Highlight
End Function
Public Property Get Left() As Double
    Left = myFrm.Left
End Property
Public Property Let Left(RHS As Double)
    myFrm.Left = RHS
End Property
Public Property Get Top() As Double
    Top = myFrm.Top
End Property
Public Property Let Top(RHS As Double)
    myFrm.Top = RHS
End Property
Public Property Get width() As Double
    If IsInited = False Then width = 0 Else width = myFrm.width
End Property
Public Property Let width(RHS As Double)
    If IsInited = False Then Exit Property
    myFrm.width = RHS
End Property
Public Property Get height() As Double
    If IsInited = False Then height = 0 Else height = myFrm.height
End Property
Public Property Let height(RHS As Double)
    If IsInited = False Then Exit Property
    myFrm.height = RHS
End Property
Private Sub myLoadLibs()
    Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Global = True
        .IgnoreCase = False
        .Multiline = True
    End With
    LibPtr = GetModuleHandle("Msftedit.dll")
    If LibPtr = 0 Then LibPtr = LoadLibrary("Msftedit.dll")
    TW2PX = GetDPI() / 72
End Sub
Private Function GetDPI() As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Call ReleaseDC(0, hdc)
End Function
Public Sub Init(ByRef TargetBox As MSForms.TextBox, Optional ByVal CodeType As String = "js")
On Error GoTo InitErr
    Set PtnDic = GetSyntaxDic(CodeType)
    If IsInited = True Then Call ClearControls
    Set myTbx = TargetBox
    myTbx.Visible = False
    Dim pControl As Variant: Set pControl = myTbx.Parent
    Set myCtr = pControl.Controls.Add("Forms.Frame.1")
    Set myFrm = myCtr
    With myCtr
        .Left = myTbx.Left
        .width = myTbx.width
        .Top = myTbx.Top
        .height = myTbx.height
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BackColor = MY_BACK_COLOR
        fHw = .[_GethWnd]
        rHw = SummonEditor
        IsInited = True
        If RichHw = 0 Then err.Raise vbObjectError + 1, , "CreateWindowEx failed"
        Me.RichText = myTbx.text
    End With
    Call Highlight
    Exit Sub
InitErr:
    Debug.Print "Init error: " & err.Number & " " & err.Description
    On Error Resume Next
    err.Raise err.Number, err.source, err.Description
End Sub
Private Sub ClearControls()
On Error Resume Next
    DestroyWindow rHw: rHw = 0
    Set myFrm = Nothing: Set myTbx = Nothing: Set myCtr = Nothing
End Sub
Private Sub Class_Terminate()
On Error Resume Next
    Call ClearControls
    Set PtnDic = Nothing: Set Reg = Nothing
End Sub
Private Sub Class_Initialize()
    IsInited = False
    MY_FONT_COLOR = RGB(220, 220, 220)
    MY_BACK_COLOR = RGB(30, 30, 30)
    Dim bytes() As Byte, i
    bytes = "Cascadia Code"
    With CF2
        .cbSize = LenB(CF2)
        .dwMask = MY_MASK
        .crTextColor = MY_FONT_COLOR
        For i = 0 To UBound(bytes)
            CF2.szFaceName(i) = bytes(i)
        Next i
    End With
    Call myLoadLibs
End Sub
Public Property Get GetSyntaxDic(CodeType As String) As Object
    Set GetSyntaxDic = CreateObject("Scripting.Dictionary")
    CodeType = StrConv(CodeType, vbLowerCase)
    With GetSyntaxDic
        If CodeType Like "js" Then
            .Add """[^""]*""|'[^']*'|`[^`]*`", RGB(214, 157, 133)
            .Add "\$\{[^}]+\}", RGB(220, 220, 170)
            .Add "\b(const|let|var|function|class|extends|super|new|import|export|from)\b", RGB(86, 156, 214)
            .Add "\b(if|else|switch|case|break|continue|for|while|do|return|try|catch|finally|throw)\b", RGB(86, 156, 214)
            .Add "\b(delete|in|instanceof|typeof|void|yield|await)\b", RGB(86, 156, 214)
            .Add "\b(0[xX][0-9a-fA-F]+|0[bB][01]+|0[oO][0-7]+|\d+(\.\d+)?([eE][+-]?\d+)?)\b", RGB(181, 206, 168)
            .Add "//[^\n]*", RGB(106, 153, 85)
            .Add "/\*[\s\S]*?\*/", RGB(106, 153, 85)
            .Add "/\/(?!\s)(?:\\.|[^/\\])+\/[gimsuy]*/", RGB(214, 157, 133)
            .Add "\b[a-zA-Z_]\w*(?=\s*:)\b", RGB(156, 220, 254)
            .Add "\b(Array|Object|String|Number|Boolean|Promise|Symbol|Map|Set|WeakMap|WeakSet|Date|Math|JSON|RegExp|Error|BigInt)\b", RGB(78, 201, 176)
            .Add "\b(map|filter|reduce|forEach|push|pop|shift|unshift|includes|startsWith|endsWith|test|exec)\b", RGB(220, 220, 170)
        ElseIf CodeType Like "cs" Then
            .Add """[^""]*""|@""[^""]*""", RGB(214, 157, 133)
            .Add "\b(class|namespace|using|public|private|protected|internal|static|void|int|string|bool|var|new|return|if|else|switch|case|break|continue|for|foreach|while|do|try|catch|finally|throw|this|base|enum|struct|interface|async|await)\b", RGB(86, 156, 214)
            .Add "\b\d+(\.\d+)?\b", RGB(181, 206, 168)
            .Add "//.*$", RGB(106, 153, 85)
            .Add "/\*[\s\S]*?\*/", RGB(106, 153, 85)
        ElseIf CodeType Like "vb" Then
            .Add """[^""]*""", RGB(214, 157, 133)
            .Add "\b(Sub|Function|End|If|Then|Else|ElseIf|For|Each|Next|Do|Loop|While|Wend|Select|Case|With|Set|Dim|As|New|Call|ByVal|ByRef|Optional|Private|Public|Const|Exit|Not|And|Or|Mod|True|False|Nothing)\b", RGB(86, 156, 214)
            .Add "\b\d+(\.\d+)?\b", RGB(181, 206, 168)
            .Add "'[^\n]*", RGB(106, 153, 85)
        Else
            .Add "(r|u|f|fr|rf)?(""" & "[\s\S]*?" & """|'''[\s\S]*?'''|""[^""\n]*""|'[^'\n]*')", RGB(214, 157, 133)
            .Add "\{[^}]+\}", RGB(220, 220, 170)
            .Add "\b(def|class|lambda|yield|global|nonlocal)\b", RGB(86, 156, 214)
            .Add "\b(if|elif|else|for|while|break|continue|return|pass)\b", RGB(86, 156, 214)
            .Add "\b(try|except|finally|raise|assert)\b", RGB(86, 156, 214)
            .Add "\b(import|from|as)\b", RGB(86, 156, 214)
            .Add "\b(and|or|not|in|is)\b", RGB(86, 156, 214)
            .Add "\b(0[xX][0-9a-fA-F]+|0[bB][01]+|0[oO][0-7]+|\d[\d_]*(\.\d[\d_]*)?([eE][+-]?\d+)?)\b", RGB(181, 206, 168)
            .Add "#.*$", RGB(106, 153, 85)
            .Add "#\s*(TODO|FIXME|NOTE)\b.*", RGB(255, 200, 130)
            .Add "@\w+", RGB(197, 134, 192)
            .Add "\b(print|len|range|open|enumerate|zip|map|filter|sum|min|max|sorted|type|isinstance|dir|vars|help)\b", RGB(78, 201, 176)
            .Add "\b(int|float|str|list|dict|set|tuple|bool|bytes|object|Exception|ValueError|TypeError|KeyError|IndexError|RuntimeError)\b", RGB(78, 201, 176)
        End If
    End With
End Function

