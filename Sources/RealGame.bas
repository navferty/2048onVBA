Attribute VB_Name = "RealGame"
Option Explicit

Public ROWS_COUNT As Integer
Public COL_COUNT As Integer
Public NEW_TILES_COUNT As Integer


Public Const START_R As Integer = 2
Public Const START_C As Integer = 2


'ROWS_COUNT COL_COUNT START_R START_C

Public PrevStates As Collection


Sub DeclareKeys()
Application.OnKey "{UP}", "MoveUp"
Application.OnKey "{DOWN}", "MoveDown"
Application.OnKey "{LEFT}", "MoveLeft"
Application.OnKey "{RIGHT}", "MoveRight"
Application.OnKey "{BACKSPACE}", "SetPrevState"
End Sub

Sub UndeclareKeys()
Application.OnKey "{UP}"
Application.OnKey "{DOWN}"
Application.OnKey "{LEFT}"
Application.OnKey "{RIGHT}"
Application.OnKey "{BACKSPACE}"
End Sub

Sub ClearCells()
ReadState
Application.EnableEvents = False
With ActiveSheet.Range(Cells(START_R, START_C), Cells(START_R + ROWS_COUNT - 1, START_C + COL_COUNT - 1))
    .ClearContents
    .Interior.Color = xlNone
End With
Application.EnableEvents = True
End Sub

Private Sub ReadState()

Dim d As New Dictionary
Dim i As Integer, j As Integer
For i = START_R To START_R + ROWS_COUNT - 1
    For j = START_C To START_C + COL_COUNT - 1
        With ActiveSheet.Cells(i, j)
            d.Add .Address(False, False), .Value
        End With
    Next j
Next i

'Set PrevState = d

If PrevStates Is Nothing Then Set PrevStates = New Collection

PrevStates.Add d
End Sub

Sub SetPrevState()

Dim i As Integer, j As Integer
Dim v As Variant
Dim PrevState As Dictionary

Application.EnableEvents = False

If PrevStates.Count > 0 Then
    Set PrevState = PrevStates.Item(PrevStates.Count)
    For Each v In PrevState
        ActiveSheet.Range(v).Value = PrevState.Item(v)
    Next
    PrevStates.Remove (PrevStates.Count)
    
    RefreshColors
End If

Application.EnableEvents = True

End Sub

Private Sub RefreshColors()

Dim i As Integer, j As Integer

Application.EnableEvents = False

For i = START_R To START_R + ROWS_COUNT - 1
    For j = START_C To START_C + COL_COUNT - 1
        With ActiveSheet.Cells(i, j)
            If .Value <> "" Then
                .Interior.Color = 100000 - 100 * .Value
            Else
                .Interior.Color = xlNone
            End If
        End With
    Next j
Next i

Application.EnableEvents = True

End Sub

Private Sub NextMove()
Dim i As Integer, j As Integer
Dim n As Integer
Dim c As New Collection
Dim cIndex As Integer
Dim cValue As Integer
Set c = GetFreeCells

Application.EnableEvents = False

If c.Count >= NEW_TILES_COUNT Then
    For i = 1 To NEW_TILES_COUNT
        Randomize
        cIndex = (c.Count - 1) * Rnd + 1 ' get one random cell from all free cells
        Randomize
        cValue = 2 * (1 + CInt(Rnd * 0.55)) '2 * CInt(Rnd) + 2
        c(cIndex).Value = cValue
        c(cIndex).Interior.Color = 320000
        c.Remove cIndex
    Next i
Else
    MsgBox "No more space"
End If

Application.EnableEvents = True

End Sub

Sub testtest()
Dim i As Integer
Dim s As Integer
'методом подбора вероятность 0.9
For i = 1 To 1000
    Randomize
    s = s + CInt(Rnd * 0.55)
    'Debug.Print CInt(Rnd * 0.55)
Next
Debug.Print s
End Sub

Private Function GetFreeCells() As Collection
Dim i As Integer, j As Integer
Dim c As New Collection
For i = START_R To START_R + ROWS_COUNT - 1
    For j = START_C To START_C + COL_COUNT - 1
        If ActiveSheet.Cells(i, j) = "" Then
            c.Add ActiveSheet.Cells(i, j)
            ActiveSheet.Cells(i, j).Interior.Color = xlNone
        Else
            ActiveSheet.Cells(i, j).Interior.Color = 100000 - 100 * ActiveSheet.Cells(i, j).Value
        End If
    Next j
Next i
Set GetFreeCells = c
End Function

Public Sub MoveLeft()
Dim i As Integer
ReadState
For i = START_R To START_R + ROWS_COUNT - 1
    MoveRow i, True
Next
NextMove
End Sub

Public Sub MoveRight()
Dim i As Integer
ReadState
For i = START_R To START_R + ROWS_COUNT - 1
    MoveRow i, False
Next
NextMove
End Sub

Public Sub MoveUp()
Dim i As Integer
ReadState
For i = START_C To START_C + COL_COUNT - 1
    MoveColumn i, True
Next
NextMove
End Sub

Public Sub MoveDown()
Dim i As Integer
ReadState
For i = START_C To START_C + COL_COUNT - 1
    MoveColumn i, False
Next
NextMove
End Sub

Private Sub MoveColumn(colIndex As Integer, isToUp As Boolean)
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim cStep As Integer
Dim ws As Worksheet
Set ws = ActiveSheet

If isToUp Then
    i = START_R
    n = ROWS_COUNT + START_R - 1
    cStep = 1
Else
    i = ROWS_COUNT + START_R - 1
    n = START_R
    cStep = -1
End If

Application.EnableEvents = False

Do While Not (i = n)
    If ws.Cells(i, colIndex).Value = ws.Cells(i + cStep, colIndex).Value And ws.Cells(i, colIndex).Value <> "" Then
        ws.Cells(i, colIndex).Value = ws.Cells(i, colIndex).Value + ws.Cells(i + cStep, colIndex).Value
        ws.Cells(i + cStep, colIndex).Value = ""
        For j = i + cStep To n Step cStep
            ws.Cells(j, colIndex).Value = ws.Cells(j + cStep, colIndex).Value
        Next j
    ElseIf ws.Cells(i, colIndex).Value = "" And IsAnythingElseColumn(ws, colIndex, i, isToUp) Then
        For j = i To n Step cStep
            ws.Cells(j, colIndex).Value = ws.Cells(j + cStep, colIndex).Value
        Next j
    Else
        i = i + cStep
    End If
Loop

Application.EnableEvents = True

End Sub

Private Sub MoveRow(rowIndex As Integer, isToLeft As Boolean)
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim cStep As Integer
Dim ws As Worksheet
Set ws = ActiveSheet

If isToLeft Then
    i = START_C
    n = COL_COUNT + START_C - 1
    cStep = 1
Else
    i = COL_COUNT + START_C - 1
    n = START_C
    cStep = -1
End If

Application.EnableEvents = False
Do While Not (i = n)
    If ws.Cells(rowIndex, i).Value = ws.Cells(rowIndex, i + cStep).Value And ws.Cells(rowIndex, i).Value <> "" Then
        ws.Cells(rowIndex, i).Value = ws.Cells(rowIndex, i).Value + ws.Cells(rowIndex, i + cStep).Value
        ws.Cells(rowIndex, i + cStep).Value = ""
        For j = i + cStep To n Step cStep
            ws.Cells(rowIndex, j).Value = ws.Cells(rowIndex, j + cStep).Value
        Next j
    ElseIf ws.Cells(rowIndex, i).Value = "" And IsAnythingElseRow(ws, rowIndex, i, isToLeft) Then
        For j = i To n Step cStep
            ws.Cells(rowIndex, j).Value = ws.Cells(rowIndex, j + cStep).Value
        Next j
    Else
        i = i + cStep
    End If
Loop

Application.EnableEvents = True

End Sub

Private Function IsAnythingElseColumn(ws As Worksheet, colIndex As Integer, rrow As Integer, isToDown As Boolean) As Boolean
Dim i As Integer, n As Integer
Dim cStep As Integer

If isToDown Then
    n = ROWS_COUNT + START_R - 1
    cStep = 1
Else
    n = START_R
    cStep = -1
End If

For i = rrow To n Step cStep
    If ws.Cells(i, colIndex) <> "" Then
        IsAnythingElseColumn = True
        Exit Function
    End If
Next i
IsAnythingElseColumn = False
End Function

Private Function IsAnythingElseRow(ws As Worksheet, rowIndex As Integer, col As Integer, isToRight As Boolean) As Boolean
Dim i As Integer, n As Integer
Dim cStep As Integer

If isToRight Then
    n = COL_COUNT + START_C - 1
    cStep = 1
Else
    n = START_C
    cStep = -1
End If

For i = col To n Step cStep
    If ws.Cells(rowIndex, i) <> "" Then
        IsAnythingElseRow = True
        Exit Function
    End If
Next i
IsAnythingElseRow = False
End Function
