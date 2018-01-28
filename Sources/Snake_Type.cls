VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private pEupId As String
Private pTail As Collection
Private pHead As Range
'Property Get EupId() As String
'    EupId = pEupId
'End Property
'Property Let EupId(EupId As String)
'    pEupId = EupId
'End Property

Private Sub Class_Initialize()
    Set pTail = New Collection
    Set pHead = ThisWorkbook.ActiveSheet.Range("AA20")
    pTail.Add pHead
End Sub

Private Sub Class_Terminate()
    Set pTail = Nothing
    Set pHead = Nothing
End Sub

Public Sub MoveNext(direction As String)
    Dim rOffset As Integer, cOffset As Integer
    
    Select Case direction
        Case "U"
            rOffset = -1
        Case "D"
            rOffset = 1
        Case "L"
            cOffset = -1
        Case "R"
            cOffset = 1
    End Select
    
    Set pHead = pHead.Offset(rOffset, cOffset)
    
    Select Case pHead.Interior.Color
        Case vbGreen
            pTail.Add pHead
            AddFood
            
        Case RGB(208, 206, 206)
            pTail.Add pHead
            pTail(1).Interior.Color = RGB(208, 206, 206) ' clear tail
            pTail.Remove 1
            
        Case Else
            MsgBox "Ooops. Score: " & pTail.Count
            ClearField
            End
    End Select
    
    pHead.Interior.Color = vbBlack
    
End Sub
