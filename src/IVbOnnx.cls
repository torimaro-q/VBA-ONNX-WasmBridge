VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbOnnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Property Get name() As String
End Property
Public Property Get Info() As String
End Property
Public Property Get JsCode() As String
End Property
Public Property Get exLibs() As Collection
End Property
Public Sub Render(ByRef GLF As GLFrame, ByRef Results As Collection, Optional ByVal imageAspect As Double = 1, Optional ByVal imageScale As Double = 1)
End Sub
Public Function Export(ByRef target As Worksheet, ByRef Parent As VbOnnxMain, Optional Left As Double = 0, Optional Top As Double = 0) As ChartObject
End Function
