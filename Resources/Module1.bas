Attribute VB_Name = "Module1"
Option Explicit

Sub StockInfo()
    Dim iRows As Double
    Dim symbol As String
    Dim ws As Worksheet
    Dim i As Double
    Dim istart As Double
    Dim arySymbol As Variant
    Dim tmpSymbol As String
    Dim openVal As Double
    Dim closeVal As Double
    Dim grtpctincr As Double
    Dim grtpctdecr As Double
    Dim grttotvol As Double
    Dim arygrtpctincr As Variant
    Dim arygrtpctdecr As Variant
    Dim arygrttotvol As Variant
    Dim aryresult As Variant
    Dim tmppct As Double
    Dim tmpvol As Double
    Dim outputRow As Integer
    
    Set ws = ActiveSheet
    iRows = ws.UsedRange.Rows.Count
    
    For Each ws In ActiveWorkbook.Worksheets
        istart = 2
        outputRow = istart
        For i = istart To iRows
            symbol = Cells(i, 1).Value
            If i = istart Then
                arySymbol = Array(symbol, i)
            Else
                If arySymbol(0) <> symbol Then
                    tmpSymbol = CStr(arySymbol(0))
                    openVal = arySymbol(1)
                    closeVal = i - 1
                    aryresult = ProcessSymbol(tmpSymbol, openVal, closeVal, outputRow, ws)
                    
                    tmppct = CDbl(Left(aryresult(0), Len(aryresult(0)) - 1))
                    tmpvol = CDbl(aryresult(1))
                    
                    If tmppct > grtpctincr Then
                        grtpctincr = tmppct
                        arygrtpctincr = Array(tmpSymbol, tmppct)
                    End If
                    If tmppct < grtpctdecr Then
                        grtpctdecr = tmppct
                        arygrtpctdecr = Array(tmpSymbol, tmppct)
                    End If
                    If tmpvol > grttotvol Then
                        grttotvol = tmpvol
                        arygrttotvol = Array(tmpSymbol, tmpvol)
                    End If
                    
                    arySymbol = Array(symbol, i)
                    outputRow = outputRow + 1
                End If
            End If
            DoEvents
        Next i
        
        Call OutputGreatest(arygrtpctincr, arygrtpctdecr, arygrttotvol, ws)
    Next
    
    MsgBox "Process Complete!"
End Sub
Sub OutputGreatest(grtpctincr As Variant, grtpctdecr As Variant, grttotvol As Variant, ws As Worksheet)
    Dim strIncr As String
    Dim dblIncr As Double
    Dim strDecr As String
    Dim dblDecr As Double
    Dim strVol As String
    Dim dblVol As Double
    Dim rng As Range
    
    strIncr = grtpctincr(0)
    dblIncr = grtpctincr(1)
    strDecr = grtpctdecr(0)
    dblDecr = grtpctdecr(1)
    strVol = grttotvol(0)
    dblVol = grttotvol(1)
    
    Set rng = ws.Range("P2")
    rng.Value = strIncr
    Set rng = ws.Range("Q2")
    rng.Value = CStr(dblIncr) + "%"
    
    Set rng = ws.Range("P3")
    rng.Value = strDecr
    Set rng = ws.Range("Q3")
    rng.Value = CStr(dblDecr) + "%"

    Set rng = ws.Range("P4")
    rng.Value = strVol
    Set rng = ws.Range("Q4")
    rng.Value = CStr(dblVol)
End Sub

Function ProcessSymbol(symbol As String, firstIdx As Double, lastIdx As Double, outputRow As Integer, ws As Worksheet) As Variant
    Dim yrchange As String
    Dim yrPercent As String
    Dim totalVol As String
    Dim openVal As Double
    Dim closeVal As Double
    Dim colYrVals As Collection

    openVal = ws.Cells(firstIdx, 3).Value
    closeVal = ws.Cells(lastIdx, 6).Value
    
    yrchange = GetYearlyChange(openVal, closeVal)
    yrPercent = GetYearlyPercent(openVal, closeVal)
    totalVol = GetTotalVolume(firstIdx, lastIdx, ws)
    Call OutputSymbolInfo(outputRow, symbol, yrchange, yrPercent, totalVol, ws)
    ProcessSymbol = Array(yrPercent, totalVol)
End Function

Sub OutputSymbolInfo(outputRow As Integer, symbol As String, yrchange As String, percentchange As String, totalvolume As String, ws As Worksheet)
    Dim rng As Range
    Dim row As String
    row = CStr(outputRow)
    
    Set rng = ws.Range("I" + row)
    rng.Value = symbol
    
    Set rng = ws.Range("J" + row)
    If Left(yrchange, 1) = "-" Then
        rng.Interior.Color = VBA.ColorConstants.vbRed
    Else
        rng.Interior.Color = VBA.ColorConstants.vbGreen
    End If
    rng.Value = yrchange
    
    Set rng = ws.Range("K" + row)
    rng.Value = percentchange
    
    Set rng = ws.Range("L" + row)
    rng.Value = totalvolume
    
End Sub

Function GetTotalVolume(firstRow As Double, lastRow As Double, ws As Worksheet) As String
    Dim rng As Range
    Dim firstCell As String
    Dim lastCell As String
    Dim total As Double
    
    firstCell = "G" + CStr(firstRow)
    lastCell = "G" + CStr(lastRow)
    
    Set rng = Range(firstCell, lastCell)
    total = WorksheetFunction.Sum(rng)
    GetTotalVolume = CStr(total)
End Function

Function GetYearlyPercent(openVal As Double, closeVal As Double) As String
    Dim val As Double
    val = ((closeVal - openVal) / openVal)
    GetYearlyPercent = Format(val * 100, "0.00") + "%"
End Function

Function GetYearlyChange(openVal As Double, closeVal As Double) As String
    Dim val As Double
    val = closeVal - openVal
    GetYearlyChange = Format(val, "0.00")
End Function

Function GetYearlyVals(ary As Variant, ws As Worksheet) As Collection
    Dim vals As Collection
    Dim openVal As Double
    Dim closeVal As Double
    Set vals = New Collection
    
    openVal = ws.Cells(ary(1), 3).Value
    closeVal = ws.Cells(ary(2), 6).Value
    
    vals.Add openVal
    vals.Add closeVal
    Set GetYearlyVals = vals
End Function





