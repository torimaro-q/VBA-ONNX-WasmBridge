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
        Dim dict As Object: Set dict = .GetModelDict
        Dim k, i
        For Each k In dict.keys()
            i = i + 1
            Set .OnnxModel = dict.Item(k)
            .Execute fpath
            With .Export(ActiveSheet)
                .Left = tX
                .Top = tY
                .width = .width * 0.8
                .height = .height * 0.8
                tX = tX + .width + 5
                If i Mod 3 = 0 Then
                    tX = 100
                    tY = tY + .height + 5
                End If
            End With
        Next k
    End With
err:
End Sub


