'Attribute VB_Name = "Módulo1"
Sub VBAchallenge()
Timestart = Timer
Dim WS As Worksheet

For Each WS In ActiveWorkbook.Sheets
WS.Select
TAREA2
COLORES
Next WS

MsgBox ("TIEMPO DE EJECUCION: " & (Timer - Timestart) & " SEGUNDOS")
End Sub


Sub TAREA2()

Fecha = ActiveSheet.Name
Range("J1").Value = "TICKER"
Range("K1").Value = "YEARLY CHANGE"
Range("L1").Value = "% CHANGE"
Range("M1").Value = "TOTAL STOCK VOLUME"
Range("Q1").Value = "TICKER"
Range("R1").Value = "VALUE"
Range("P2").Value = "GREATEST %INCREASE"
Range("P3").Value = "GREATEST %DECREASE"
Range("P4").Value = "GREATEST TOTAL VOLUME"



endRow = Range("A2").End(xlDown).Row
Range("J2").Value = Range("A2").Value
j = 2

For i = 3 To endRow
If Range("A" & i).Value = Range("J" & j).Value Then
j = j
Else
Range("J" & j + 1).Value = Range("A" & i).Value
j = j + 1
End If
Next i
    
    
g = 2
m = 2
suma = 0
    For i = g To endRow
    If Range("A" & i).Value = Range("J" & m).Value Then
        suma = suma + Range("G" & i)
        If Range("B" & i).Value = Fecha & "0102" Then
        ipx = Range("C" & i).Value
        ElseIf Range("B" & i).Value = Fecha & "1231" Then
        cpx = Range("F" & i).Value
        g = Range("C" & i).Row + 1
        Range("K" & m).Value = cpx - ipx
        Range("L" & m).Value = ((cpx - ipx) / ipx)
        Range("L" & m).Value = FormatPercent(Range("L" & m).Value)
        Range("M" & m).Value = suma
        m = m + 1
        suma = 0
        End If
    ElseIf Range("A" & i).Value <> Range("J" & m).Value Then
    cpx = 0
    ipx = 0
    End If
Next i
    
'GREAT INCR
endS = Range("J2").End(xlDown).Row
greatIn = Range("L2")
tickerr = Range("J2").Value
greatDec = Range("L2")
tickerr2 = Range("J2").Value
greatTot = Range("M2")
tickerr3 = Range("J2").Value


For q = 3 To endS
    If Range("L" & q) > greatIn Then
    greatIn = Range("L" & q)
    tickerr = Range("J" & q).Value
    ElseIf Range("L" & q) < greatDec Then
    greatDec = Range("L" & q)
    tickerr2 = Range("J" & q).Value
    ElseIf Range("M" & q) > greatTot Then
    greatTot = Range("M" & q)
    tickerr3 = Range("J" & q).Value
    Else
    greatIn = greatIn
    tickerr = tickerr
    greatDec = greatDec
    tickerr2 = tickerr2
    greatTot = greatTot
    tickerr3 = tickerr3
    End If
Next q
Range("R2") = FormatPercent(greatIn, 2)
Range("Q2") = tickerr
Range("R3") = FormatPercent(greatDec, 2)
Range("Q3") = tickerr2
Range("R4") = greatTot
Range("Q4") = tickerr3




End Sub
Sub COLORES()
'
' COLORES Macro
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65280
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65280
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub



