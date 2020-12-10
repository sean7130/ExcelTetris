Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub Example()
     Dim StartTime As Currency
     Dim EndTime As Currency
     Dim TickFrequency As Currency
     
     QueryPerformanceFrequency TickFrequency
     
     QueryPerformanceCounter StartTime
     QueryPerformanceCounter EndTime
     
     MsgBox TickFrequency
     
     MsgBox "Elapsed time: " & 1000 * (EndTime - StartTime) / TickFrequency & "ms" 'scaled to milliseconds
End Sub


Sub Main()
'
' Macro1 Macro
'

'
    Cells.Select
    Selection.ColumnWidth = 2.14
    
    Call board_demostrator
    
    Dim i As Integer
    Dim diff As Long
    Dim EndTime As Currency
    Dim StartTime As Currency
    Dim TickFrequency As Currency

    QueryPerformanceFrequency TickFrequency
    
    i = 0
    Dim dead As Boolean
    dead = False
    While Not dead
        QueryPerformanceCounter StartTime
        QueryPerformanceCounter EndTime
        'MsgBox (EndTime - StartTime) / TickFrequency
    
        i = i + 1
        If i = 10 Then
            dead = True
            MsgBox i
        End If
    Wend
    
End Sub

Sub reset_all()
    Call zero_data
    Call update_front_from_data
End Sub

Sub zero_data()
    Sheets("data").Range(Sheets("data").Cells(2, "b"), Sheets("data").Cells(21, "k")).Value = 0
End Sub

Sub update_front_from_data()
    Dim i As Byte
    Dim j As Byte
    Dim this_val As Byte
    
    For i = 2 To 21
        For j = 2 To 11
            Call updateBoardAbs(j, i, Sheets("data").Cells(i, j).Value)
        Next
    Next
    
End Sub

Sub board_demostrator()
'
' board_demostrator Macro
'

'
    Range("B2:K21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub CondFormat()
'
' CondFormat Macro
'

'
    Range("B2:K21").Select
End Sub
Function moveAllMovableDownwards()
' This function moves all the blocks that should be moved:
' Which are the blocks assoicated with datavalue == 1

    Dim i As Byte
    Dim j As Byte
    
    Dim row_idx As Byte
    Dim col_idx As Byte
    
    ' goal is to iterate backwards from 21 to 2 for row
    ' and then iterate backwards from 11 to 2 for columns
    For i = 0 To 19
        For j = 0 To 9
        row_idx = 21 - i
        col_idx = 11 - j
        If Sheets("data").Cells(row_idx, col_idx).Value = 1 Then
            Call moveBlock(col_idx, row_idx, 0, 1)
            
        End If
            
        Next
    Next


End Function

Function moveBlock(x, y, delta_x, delta_y)
    Dim val As Integer
    val = Sheets("data").Cells(y, x).Value
    Call updateBoardAbs(x, y, 0)
    Call updateBoardAbs(x + delta_x, y + delta_y, val)
End Function

Function updateBoardAbs(x, y, val)
    ' function does the follow two tasks:
    ' update visuals for the front
    ' write data at the back
    
    'val: 1 is gray
    'val: 2 is black
    'val: 0 is nothing
    
    If x < 2 Then
        MsgBox "updateBoardAbs: X cannot be smaller than 2"
        Exit Function
    ElseIf x > 11 Then
        MsgBox "updateBoardAbs: X cannot be greater than 11"
        Exit Function
    End If
    
    If y < 2 Then
        MsgBox "updateBoardAbs: Y cannot be smaller than 2"
        Exit Function
    ElseIf y > 21 Then
        MsgBox "updateBoardAbs: Y cannot be greater than 21"
        Exit Function
    End If
    
    Sheets("data").Cells(y, x).Value = val
    
    If val = 1 Then
        With Sheets("front").Cells(y, x).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.5
            .PatternTintAndShade = 0
        End With
        '
    ElseIf val = 2 Then
            With Sheets("front").Cells(y, x).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    ElseIf val = 0 Then
        With Sheets("front").Cells(y, x).Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    
    'Updates: data with the correct data
    'Sheets("data").Cells(y, x).Value = "testing"
    

End Function
Sub GraySelection()
'
' GraySelection Macro
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
End Sub
Sub setupSheet()
    ActiveSheet.Name = "front"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Name = "testing"
End Sub

Sub b2k21()
'
' b2k21 Macro
'

'
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:K2"), Type:=xlFillDefault
    Range("B2:K2").Select
    Range("B2:K2").Select
    Selection.AutoFill Destination:=Range("B2:K21"), Type:=xlFillDefault
    Range("B2:K21").Select
    Range("L21").Select
End Sub

