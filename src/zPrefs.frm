VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zPrefs 
   Caption         =   "VbOnnxPrefs"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "zPrefs.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "zPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Property Get SeverCode() As String
    SeverCode = Me.TextBox1.Value
End Property
