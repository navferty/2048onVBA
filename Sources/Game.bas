Attribute VB_Name = "Game"
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)

Sub Play2048()

Dim ws As Worksheet
Set ws = ActiveSheet
Dim i As Integer
Dim n As Integer
Dim val As Integer

n = 16

If ws.Cells(1, n).Value <> "" Then
    ActiveSheet.Rows(1).ClearContents
End If

Do While ws.Cells(1, n).Value = ""
    For i = 1 To n
        If ws.Cells(1, i).Value = "" Then
            ws.Cells(1, i).Value = 2 'GetValue
            Exit For
        End If
    Next i
    For i = 1 To n
        If ws.Cells(1, i).Value = ws.Cells(1, i + 1).Value And ws.Cells(1, i).Value <> "" Then
            ws.Cells(1, i).Value = ws.Cells(1, i).Value + ws.Cells(1, i + 1).Value
            ws.Columns(i + 1).EntireColumn.Delete
        End If
    Next i
    Sleep 1
    DoEvents
Loop
MsgBox "—чет " & Application.Sum(ws.Rows(1))
End Sub

Private Function GetValue() As Integer
Dim v As Integer, i As Integer
Randomize
For i = 1 To 10
    v = v + Rnd()
Next

If v > 8 Then
    GetValue = 4
Else
    GetValue = 2
End If

End Function


Sub StopPlay()
End
End Sub

Sub StopAndCleanSheet()
ActiveSheet.Rows(1).ClearContents
End
End Sub
