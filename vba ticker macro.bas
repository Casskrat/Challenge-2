Attribute VB_Name = "Module1"
Sub ticker()

Dim rowcount As Long
Dim tick As String
Dim ychange As Double
Dim pchange As Double
Dim volume As LongPtr
Dim ws As Worksheet
Dim sumrow As Long
Dim count As Integer
Dim greatest As Double
Dim smallest As Double
Dim volume2 As LongPtr


countsheet = ActiveWorkbook.Worksheets.count

For Each ws In Worksheets
    ws.Activate
    sumrow = 2
    rowcount = Range("A1").End(xlDown).Row
    count = 0

    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"

    For j = 2 To rowcount
        If Cells(j + 1, 1).Value = Cells(j, 1).Value Then
            volume = volume + Cells(j, 7).Value
            count = count + 1
        Else
            tick = Cells(j, 1).Value
            volume = volume + Cells(j, 7).Value
            Range("J" & sumrow).Value = tick
            Range("M" & sumrow).Value = volume
        
            ychange = Cells(j, 6).Value - Cells(j - count, 3).Value
            Range("K" & sumrow).Value = ychange
            pchange = Cells(sumrow, 11).Value / Cells(j - count, 3).Value
            Range("L" & sumrow).Value = pchange
        
                If ychange < 0 Then
                    Range("K" & sumrow).Interior.ColorIndex = 3
                Else
                    Range("K" & sumrow).Interior.ColorIndex = 4
                End If
                
                If pchange < 0 Then
                    Range("L" & sumrow).Interior.ColorIndex = 3
                Else
                    Range("L" & sumrow).Interior.ColorIndex = 4
                End If
        
            sumrow = sumrow + 1
            volume = 0
            count = 0
        End If
    Next j

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    greatest = WorksheetFunction.Max(Range("L2:L700000"))
    smallest = WorksheetFunction.Min(Range("L2:L700000"))
    volume2 = WorksheetFunction.Max(Range("M2:M700000"))

    For k = 2 To rowcount
        If (Cells(k, 12).Value = greatest) Then
            Cells(2, 16).Value = Cells(k, 10).Value
        ElseIf (Cells(k, 12).Value = smallest) Then
            Cells(3, 16).Value = Cells(k, 10).Value
        ElseIf (Cells(k, 13).Value = volume2) Then
            Cells(4, 16).Value = Cells(k, 10).Value
        End If

    Next k

    Range("Q2").Value = greatest
    Range("Q3").Value = smallest
    Range("Q4").Value = volume2
Next

End Sub
















