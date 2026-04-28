Attribute VB_Name = "Sample"
Option Explicit
Sub Sample1()
    OnnxUI.Show
End Sub
Sub Sample2()
    Dim fpath As String, oxm As VbOnnxMain
    Set oxm = New VbOnnxMain
    On Error GoTo err:
    With oxm
        Set .OnnxModel = New VbOnnxYOLOX
        .EnsureRuntimeFiles
        If .ExecuteFlg Then
            fpath = Application.GetOpenFilename(.FILE_FILTER)
            If Not (fpath Like "*.*") Then Exit Sub
            .Execute fpath
            With .Export(ActiveSheet)
                .Left = 100
            End With
        Else
            If MsgBox("You need to download the ONNX model and place the file in the specified folder.", vbOKCancel) = vbOK Then
                .OpenTempFolder
                .OpenEdgeForDownload
            End If
        End If
    End With
err:
End Sub

Sub Sample3()
    Dim fpath As String, oxm As VbOnnxMain, tX As Double, tY As Double
    Set oxm = New VbOnnxMain
    On Error GoTo err:
    tX = 100
    tY = 0
    With oxm
        .EnsureRuntimeFiles
        fpath = Application.GetOpenFilename(.FILE_FILTER)
        Dim names As Variant: names = .ModelDict.keys()
        Dim k, i, j
        For Each k In names
            i = i + 1
            Set .OnnxModel = .GetModelFromName(CStr(k))
            .Execute fpath
            With .Export(ActiveSheet)
                .Left = tX
                .Top = tY
                .width = .width * 0.8
                .height = .height * 0.8
                tX = tX + .width + 5
                If i Mod 4 = 0 Then
                    tX = 100
                    tY = tY + .height + 5
                End If
            End With
            For j = 1 To 10
                DoEvents
                Sleep 100
                DoEvents
            Next j
        Next k
    End With
err:
End Sub


