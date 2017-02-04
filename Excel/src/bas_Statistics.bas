Attribute VB_Name = "bas_Statistics"
Option Explicit


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

                                  
Sub MultipleRegression()
     'P138
     'Code from http://www.ozgrid.com/forum/showthread.php?t=173701
     'This will only work in Excel and is not very flexible.  Normally Excel regression
     'for VBA uses the Excel functions to perform the matrix algebra, but this will not translate
     'to Access or other programs.   until the LINEST function is rebuilt at the code level.
                         
    Dim a As Range, n As Long, j As Long, k As Long
    Dim y, x, m, SeCo() As Double
    Dim Coeff, rsq As Single
    
    'not defined on first build
    Dim cname, rvar
    
    If Intersect(Range("A5").CurrentRegion, ActiveCell) Is Nothing Then _
    MsgBox "Select cell within the data range" & Chr(10) & _
    "for the regression code to run": Exit Sub
     
    With Range("A4", ActiveCell)
        n = .Rows.Count
        k = .Columns.Count
        .Columns(k).Offset(, 1).Insert xlToRight
        .Columns(k + 1) = 1
        y = .Resize(n - 1, 1).Offset(1)
        x = .Resize(n - 1, k).Offset(1, 1)
        .Offset(, k).Resize(, 1).Delete xlToLeft
        cname = Application.Transpose(.Resize(, k - 1).Offset(, 1))
    End With
    ReDim SeCo(1 To k, 1 To 1)
     
    With Application
        m = .MInverse(.MMult(.Transpose(x), x))
        Coeff = .MMult(m, .MMult(.Transpose(x), y))
        rvar = (.SumSq(y) - Evaluate(.MMult(.Transpose(y), .MMult(x, Coeff)))) / (n - k)
        rsq = Evaluate(.MMult(.Transpose(y), .MMult(x, Coeff))) / .SumSq(y)
        For j = 1 To k
            SeCo(j, 1) = (rvar * m(j, j)) ^ 0.5
        Next j
    End With
     
    With Range("M1")
        .Value = "Coeffs"
        .Offset(1).Resize(k) = Coeff
        .Offset(, 1) = "SECoef"
        .Offset(1, 1).Resize(k) = SeCo
        .Offset(, 2) = "RSq"
        .Offset(1, 2) = rsq
        .Offset(1, -1).Resize(k - 1) = cname
        .Offset(k, -1) = "Const"
    End With
End Sub

'-----------------------------------------------------------------------------------------------------------------------
'Begin calculation for Regression equation
Sub PerformRegressionEquation()
    createCalculationWorksheet
    RegressionEquation
    
End Sub


Public Function RegressionEquation()
   'Created by Bill Young
   'I am not able to figure out how to feed the formula with the ranges
   'The ranges have to be entered in the code in two locations
   'Also, inserting the new worksheet for calculations doesn't work when inside this function.
   'Not sure why this is the case
    Dim x As Variant, y As Variant, Results As String
    Dim a As Double
    Dim b As Double
    Dim denom As Double
    Dim aNum As Double
    Dim bNum As Double
    Dim xRng As Variant
    Dim yRng As Variant
    
    Dim ws As Object
    Set ws = Application.Sheets("sheet5")
    ws.Select
    
    'createCalculationWorksheet
    
    Set xRng = Range("B1:B7")
    Set xRng = Range("C1:C7")
    
    
'    If IsArray(x) _
'    Then
'        xRng = x
'    Else
'        xRng = Array(x)
'    End If
'
'    If IsArray(y) _
'    Then
'        yRng = y
'    Else
'        yRng = Array(y)
'    End If
    
    ws.Range("B1:B7").Copy _
        Destination:=Worksheets("RegEqnCalcs").Range("B1")
        
    ws.Range("C1:C7").Copy _
        Destination:=Worksheets("RegEqnCalcs").Range("C1")
    
    'xRng.Select
    'Selection.Copy

    'All regression calculations are performed on RegEqnCalcs
    'Results are on RegEqnResults

'    Sheets("RegEqnCalcs").Select
'    Range("B1").Select
'    ActiveSheet.Paste
'
'    yRng.Copy
'    Sheets("RegEqnCalcs").Select
'    Range("C1").Select
'    ActiveSheet.Paste
    
    Sheets("RegEqnCalcs").Select
    
    Range("D1").Select
    Selection.Value = "XY"
    Range("E1").Select
    Selection.Value = "x^2"
    Range("F1").Select
    Selection.Value = "y^2"
    Range("G1").Select
    Selection.Value = "Ex"
    Range("H1").Select
    Selection.Value = "Ey"
    Range("I1").Select
    Selection.Value = "n"
    Range("J1").Select
    Selection.Value = "Exy"
    Range("K1").Select
    Selection.Value = "Ex^2"
    Range("L1").Select
    Selection.Value = "Ey^2"
    
    Dim rCounter As Double
    Set ws = Worksheets("RegEqnCalcs")
        With ws
            rCounter = .Cells(.Rows.Count, "B").End(xlUp).Row
        End With
    Range("D2:D" & rCounter).Select
    
    enterFormulaXY
    enterFormulaX2
    enterFormulaY2
    enterFormulaEx
    enterFormulaEy
    enterFormulaN
    enterFormulaExy
    enterFormulaEX2
    enterFormulaEY2

    
    Dim Ex As Double
        Ex = Range("G2").Value
    Dim Ey As Double
        Ey = Range("H2").Value
    Dim n As Double
        n = Range("I2").Value
        
        
    Dim Exy As Double
        Exy = Range("J2").Value
    Dim Ex2 As Double
        Ex2 = Range("K2").Value
    Dim Ey2 As Double
        Ey2 = Range("L2").Value
        
    denom = n * Ex2 - Ex ^ 2

    Sheets.Add.Name = "RegEqnResults"
       
    a = ((Ey * Ex2) - (Ex * Exy)) / denom
    b = ((n * Exy) - (Ex * Ey)) / denom

    Sheets("RegEqnResults").Select
    Range("A1").Value = "a"
    Range("A2").Value = a
    Range("B1").Value = "b"
    Range("B2").Value = b

    
    RegressionEquation = "a = " & a & ";  b = " & b
Range("A4").Value = "y " & Round(a, 4) & " + " & Round(b, 4)
    
   End Function
Sub enterFormulaXY()
    'gives XY in coulumn D
    Dim ws As Object
    Dim rCounter As Double
    Dim i As Long
    Set ws = Worksheets("RegEqnCalcs")
        
    With ws
        rCounter = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With

    With ws
        For i = 2 To rCounter
            If Len(Trim(.Range("B" & i).Value)) <> 0 _
            Then
                .Range("D" & i).Formula = "=B" & i & "*" & "C" & i
            End If
            
        Next i
    End With
End Sub

Sub enterFormulaX2()
    'gives X^2 in coulumn D
    Dim ws As Object
    Dim rCounter As Double
    Dim i As Long
    Set ws = Worksheets("RegEqnCalcs")
        
    With ws
        rCounter = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With

    With ws
        For i = 2 To rCounter
            If Len(Trim(.Range("B" & i).Value)) <> 0 _
            Then
                .Range("E" & i).Formula = "=B" & i & "*" & "B" & i
            End If
            
        Next i
    End With
End Sub

Sub enterFormulaY2()
    'gives X^2 in coulumn D
    Dim ws As Object
    Dim rCounter As Double
    Dim i As Long
    Set ws = Worksheets("RegEqnCalcs")
        
    With ws
        rCounter = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With

    With ws
        For i = 2 To rCounter
            If Len(Trim(.Range("B" & i).Value)) <> 0 _
            Then
                .Range("F" & i).Formula = "=C" & i & "*" & "C" & i
            End If
            
        Next i
    End With
End Sub
Sub enterFormulaEx()
    Range("G2").Formula = "=sum(B:B)"
End Sub
Sub enterFormulaEy()
    Range("H2").Formula = "=sum(C:C)"
End Sub
Sub enterFormulaN()
    Range("I2").Formula = "=COUNT(C:C)"
End Sub
Sub enterFormulaExy()
    Range("J2").Formula = "=SUM(D:D)"
End Sub
Sub enterFormulaEX2()
    Range("K2").Formula = "=SUM(E:E)"
End Sub
Sub enterFormulaEY2()
    Range("L2").Formula = "=SUM(F:F)"
End Sub
Private Sub createCalculationWorksheet()
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "RegEqnCalcs"
    End With
End Sub
'End calculation for regression equation
'-----------------------------------------------------------------------------------------------------------------------


