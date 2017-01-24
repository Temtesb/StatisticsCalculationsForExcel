Attribute VB_Name = "basStatistics"
Option Explicit

    
    Public Function GeometricMean(rng As Range)
    'Created by Bill Young p38
    'Note: this calculation excludes non-numeric values like Excel's GEOMEAN function.
    'Unlike GEOMEAN, negative numbers are included.
    'To properly perform the calculation with negative numbers, the absolute value of all numbers is used.
        Dim a As Variant
        Dim x As Variant
        Dim n As Double
        Dim v As Double
        Dim vt As VbVarType
        If IsArray(rng) _
        Then
            a = rng
        Else
            a = Array(rng)
        End If
        For Each x In a
            vt = VarType(x)
            If Abs(vt) > 1 And Abs(vt) < 7 _
            Then
                If Abs(x) > 0 _
                Then
                    n = Abs(n) + 1
                    v = Abs(v) + Log(Abs(x))
                End If
            End If
        Next
        If n _
        Then
           GeometricMean = Exp(v / n)
        Else: Err.Raise 13
        End If
    End Function
    
    Public Function QuartileRange(rng As Range)
    'Created by Bill Young p46
        QuartileRange = WorksheetFunction.Quartile_Exc(rng, 3) - WorksheetFunction.Quartile_Exc(rng, 1)
    End Function
    
    Public Function CoefficientOfVariation(rng As Range)
    'Created by Bill Young p44
        Dim a As Variant
        If IsArray(rng) _
        Then
            a = rng
        Else
            a = Array(rng)
        End If
        WorksheetFunction.StDev_P (a)
        CoefficientOfVariation = Format(WorksheetFunction.StDev_P(a) / WorksheetFunction.Average(a), "Standard")
    End Function
    
    Public Function Zscore(x As Double, rng As Range)
        'Created by Bill Young p46
        Dim a As Variant
        Dim s As Double
        Dim m As Double
        If IsArray(rng) _
        Then
            a = rng
        Else
            a = Array(rng)
        End If
        m = WorksheetFunction.Average(a)
        s = WorksheetFunction.StDev_P(a)
        Zscore = Format((x - m) / s, "Standard")
    End Function
    
    Public Function CorrelationCoefficient(Xrng As Range, Yrng As Range)
        'Created by Bill Young p55
        Dim a As Variant
        Dim Rxy As Variant
        Dim x As Variant
        Dim y As Variant
        Dim Sx As Double
        Dim Sy As Double
        Dim Yx As Double
        If IsArray(Xrng) _
        Then
            x = Xrng
        Else
            x = Array(Xrng)
        End If
        
        If IsArray(Yrng) _
        Then
            y = Yrng
        Else
            y = Array(Yrng)
        End If
        Rxy = WorksheetFunction.Covariance_S(x, y)
        Sx = WorksheetFunction.StDev_S(x)
        Sy = WorksheetFunction.StDev_S(y)
        CorrelationCoefficient = Format(Rxy / (Sx * Sy), "Standard")
    End Function
    
    Sub CalculateMultiMode(rng As Range)
        'The results are on the RankWorking tab.
        Dim strResultsSheetName As String
        strResultsSheetName = "RankWorking"
        Application.ScreenUpdating = False
        'Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Dim ws As Worksheet
        Set ws = CreateWorksheet(strResultsSheetName)
        ws.Activate ' check this in Excel
        ws.Paste
        ws.Range("C1").Select
        ws.Paste
        Application.CutCopyMode = False
        ws.Range("C1").Select
        ws.Range(Selection, Selection.End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
        ws.Range("D2").Formula = "=COUNTIF(A:A,C2)"
        ws.Range("D2").Copy
        ws.Range("D2").Select
        ws.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        ws.Range("C1").Select
        ws.Application.Selection.End(xlDown).Select
        ws.Application.ActiveCell.Offset(0, 1).Select
        ws.Range(Selection, Selection.End(xlUp)).Select
        ws.Paste
        Application.CutCopyMode = False
        ws.Application.Selection.Copy
        ws.Application.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        ws.Range("D1").Value = "Count"
        ws.Columns("C:D").Select
        ws.Application.Selection.AutoFilter
        ws.AutoFilter.Sort.SortFields.Clear
        ws.AutoFilter.Sort.SortFields.Add Key:= _
            Range("D1:D6513"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal
        With ws.AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ws.Application.Selection.End(xlUp).Select
        'ThisWorkbook.Sheets("RankWorking").Delete
        
        ws.Columns("A:B").Delete
        ws.Range("A1").Select
        Application.ScreenUpdating = True
    End Sub
    
    Public Function InterQuartileRange(rng)
        Dim a
        If IsArray(rng) Then a = rng Else a = Array(rng)
        'InterQuartileRange = Quartile.exc(rng, 3) - Quartile.exc(rng, 1)
        Set ActiveCell.Formula = "= Quartile.exc(rng, 3) - Quartile.exc(rng, 1)"
    End Function

