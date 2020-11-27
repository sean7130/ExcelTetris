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
     Sleep 2
     QueryPerformanceCounter EndTime
     
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
    i = 0
    Dim dead As Boolean
    dead = False
    While Not dead
        i = i + 1
        If i = 10 Then
            dead = True
            MsgBox i
        End If
    Wend
    
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

Sub testUBA()
    Call updateBoardAbs(3, 8, 1)
    

End Sub


Function updateBoardAbs(x, y, val)
    ' function does the follow two tasks:
    ' update visuals for the front
    ' write data at the back
    If x < 2 Then
        MsgBox "updateBoardAbs: X cannot be smaller than 2"
        Exit Function
    ElseIf x > 10 Then
        MsgBox "updateBoardAbs: X cannot be greater than 10"
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

