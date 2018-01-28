Attribute VB_Name = "Snake"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)

Type KeyboardBytes
    kbb(0 To 255) As Byte
End Type
Declare PtrSafe Function GetKeyboardState Lib "User32.DLL" (kbArray As KeyboardBytes) As Long

'Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Public Const SNAKE_ROWS_COUNT As Integer = 100
Public Const SNAKE_COL_COUNT As Integer = 100
Public Const SNAKE_START_R As Integer = 2
Public Const SNAKE_START_C As Integer = 2
Public SNAKE_DIRECTION As String
Public SNAKE_IS_MOVING As Boolean

Public MySnake As Snake_Type

Public Sub StartSnake()
    Dim i As Integer
    
    If MySnake Is Nothing Then
        Set MySnake = New Snake_Type
    End If
    
    DeclareSnakeKeys
    
    For i = 1 To 5000
        AddFood
    Next i
    
    SNAKE_DIRECTION = "R"
    
    Sleep 1000
    
    ContinueSnake
    
End Sub

Public Sub StopSnake()
    ClearField
    End
End Sub

Sub ContinueSnake()

    Dim i As Integer
    
    Do
        MySnake.MoveNext SNAKE_DIRECTION
        
        For i = 1 To 15
            Sleep 2
            DoEvents
            SNAKE_DIRECTION = GetDirectionFromKeyboard()
        Next i
    Loop
End Sub

Private Function GetDirectionFromKeyboard()
    Dim kbArray As KeyboardBytes
    GetKeyboardState kbArray
    If kbArray.kbb(37) = 128 Then
        GetDirectionFromKeyboard = "L"
    ElseIf kbArray.kbb(38) = 128 Then
        GetDirectionFromKeyboard = "U"
    ElseIf kbArray.kbb(39) = 128 Then
        GetDirectionFromKeyboard = "R"
    ElseIf kbArray.kbb(40) = 128 Then
        GetDirectionFromKeyboard = "D"
    Else
        GetDirectionFromKeyboard = SNAKE_DIRECTION
    End If
End Function

Public Sub AddFood()
    Dim r As Integer
    Dim c As Integer
    
    Do
        Randomize
        r = (SNAKE_ROWS_COUNT - SNAKE_START_R) * Rnd + 1
        Randomize
        c = (SNAKE_COL_COUNT - SNAKE_START_C) * Rnd + 1
        With ThisWorkbook.ActiveSheet.Cells(r, c)
            If .Interior.Color = RGB(208, 206, 206) Then
                .Interior.Color = vbGreen
            End If
        End With
    Loop While Not ThisWorkbook.ActiveSheet.Cells(r, c).Interior.Color = vbGreen
    
End Sub
'
'Sub DirectionUp()
'    SNAKE_DIRECTION = "U"
'    ContinueSnake
'End Sub
'
'Sub DirectionDown()
'    SNAKE_DIRECTION = "D"
'    ContinueSnake
'End Sub
'
'Sub DirectionLeft()
'    SNAKE_DIRECTION = "L"
'    ContinueSnake
'End Sub
'
'Sub DirectionRight()
'    SNAKE_DIRECTION = "R"
'    ContinueSnake
'End Sub
'
'Sub TogglePause()
'    SNAKE_IS_MOVING = Not SNAKE_IS_MOVING
'End Sub
'
Sub DeclareSnakeKeys()
Application.OnKey "{UP}", ""
Application.OnKey "{DOWN}", ""
Application.OnKey "{LEFT}", ""
Application.OnKey "{RIGHT}", ""
Application.OnKey "{BACKSPACE}", ""
End Sub


Sub ClearField()
    ThisWorkbook.ActiveSheet.Range("B2:CV100").Interior.Color = RGB(208, 206, 206)
End Sub
