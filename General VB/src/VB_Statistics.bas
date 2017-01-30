Attribute VB_Name = "VB_Statistics"
Option Explicit
'Authored 2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil> and William Young
'Norfolk Naval Shipyard
     'CC0 1.0 <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...
'These functions are intendid to work without dependances, using late binding when necessary

Public Enum QuartileType
    QuartileMinVal = 0
    QuartileFirst = 1
    QuartileSecond = 2
    QuartileThird = 3
    QuartileMaxVal = 4
End Enum
    
'The first 16 Methods are detailed in the 'Journal of Statistics Education' Volume 14, number 3 (2006),
    'ww2.amstat.org/publications/jse/v14n3/langford.html
Public Enum QuartileMethod
    TodoQmHoggAndLedolter = 0
    QmInclusive = 1
    QmExclusive = 2
    'TodoQmTukey = 3 'Same as method 1
    TodoQmCdf = 4
    TodoQmM_S = 5
    TodoQmLohninger = 6
    'TodoQmVining = 7 'appears to result in the same as methods 1 and 3
    TodoQmJ_F = 8
    TodoQmH_L = 9
    TodoQmH_L2 = 10
    TodoQmMinitab = 11
    TodoQmExcelInc = 12
    TodoQmSas1 = 13
    TodoQmSas2 = 14
    TodoQmSas3 = 15
    TodoQmExcelExcl = 16
End Enum

   Public Function GeometricMean(rng As Variant)
   'Created by Bill Young p38
   'Note: this calculation excludes non-numeric values like Excel's GEOMEAN function.
   'Unlike GEOMEAN, negative numbers are included.
   'Still need to confirm that this is the same method used in the /Excel function being used
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
       For Each x In a 'each element in array
           vt = VarType(x)
           If vt + 1 > 1 And vt + 1 < 7 _
           Then
               If x + 1 > 0 _
               Then
                   n = n + 1
                   v = v + 1 + Log(x + 1)
               End If
           End If
       Next
       If n _
       Then
           GeometricMean = Exp(v / n) - 1
       Else
                   MsgBox "No non zero values exist in the supplied data set"
       End If
   End Function

Public Function IsRange(obj As Variant) As Boolean 'can be run from excel or access
    On Error Resume Next
    Dim strName As String
    Select Case True
        Case InStrRev(Application.Name, "Excel") > 0
            strName = obj.Range.Name
            IsRange = Err.Number = 0
        Case InStrRev(Application.Name, "Access") > 0
            IsRange = TypeName(obj) = "Range"
    End Select
End Function

Public Function QuartileFromObject( _
    obj As Variant, _
    quart As QuartileType, _
    qMethod As QuartileMethod, _
    Optional intDimentionRank As Integer _
) As Double
    'Plan is to accept multidimetional arrays, excel ranges, word tables, access recordsets, tabledefs, and querydef and process dedicated column for quartile
    'Using late binding so that we don't have to deal with assigning referances
    Dim oThisApplication As Object: Set oThisApplication = Application
    Dim aryTemp() As Variant
    Dim aryToProcess() As Variant
    Dim oExcel As Object
    Dim oAccess As Object
    'Assign any acceptable obect to our aryToProcess
    Select Case True
        Case IsRange(obj) And InStrRev(oThisApplication.Name, "Excel") > 0
            aryTemp = ExcelRangeToNumericSafeArray(obj)
            If ArrayRank(aryTemp) > 1 Then
                aryToProcess = ArraySingleFromMultiDimention(aryTemp, intDimentionRank)
            Else
                aryToProcess = aryTemp
            End If
        Case TypeName(obj) = "Recordset" 'And InStrRev(oThisApplication.Name, "Access") > 0
            aryToProcess = oThisApplication.GetRows
        Case TypeName((obj) = "TableDef") Or (TypeName(obj) = "QueryDef")
            aryToProcess = obj.OpenRecordset(obj.Name).GetRows
        Case Else
    End Select
    If IsArray(obj) Then
        If ArrayRank(obj) > 1 Then
            aryToProcess = ArraySingleFromMultiDimention(aryToProcess, intDimentionRank)
        End If
        QuartileFromObject = Quartile(aryToProcess, quart, qMethod)
    Else
        Err.Raise 3302, Description:="The object passed to the Quartile must be an array, excel range" ', word table, access recordset, access tabledef, or access querydef."
        Exit Function
    End If
End Function

Public Function InterQuartileRange(ByRef ary As Variant, ByRef qMethod As QuartileMethod)
    'Sort the array once here so we don't sort it twice
    SortArrayInPlace (ary)
    InterQuartileRange = Quartile(ary, QuartileThird, qMethod, True) - Quartile(ary, QuartileFirst, qMethod, True)
End Function
'Excel only code Left here so that we can test and compare speed of the excel native function, if it's significantly faster then we should use it over our custom solution
'Public Function InterQuartileRange(rng)
'    Dim a
'    If IsArray(rng) Then a = rng Else a = Array(rng)
'    'InterQuartileRange = Quartile.exc(rng, 3) - Quartile.exc(rng, 1)
'    Set ActiveCell.Formula = "= Quartile.exc(rng, 3) - Quartile.exc(rng, 1)"
'End Function


Public Function Quartile( _
                    ByRef ary As Variant, _
                    ByRef quart As QuartileType, _
                    ByRef qMethod As QuartileMethod, _
                    Optional fArrayIsPresorted As Boolean = False _
) As Double
'Created by Jeremy Gerdes to remove dependacies on MS Excel
'Sort array in ascending order
'Find median for each quartiel, Q1,Q2,Q3
'Eric Langford lists 15 Methods here:
'   https://ww2.amstat.org/publications/jse/v14n3/langford.html
'for completeness those 15 methods will be iterated in the enum of methods, but only those that have beed written (as needed) will be uncommented from the enum.
'http://stat.ethz.ch/R-manual/R-patched/library/stats/html/quantile.html
'Author(s) of the version used in R >= 2.0.0, Ivan Frohne and Rob J Hyndman.
'R programming language is an open source project.
'https://cran.r-project.org/src/base/R-3/
'View R Source code:http://stackoverflow.com/questions/19226816/how-can-i-view-the-source-code-for-a-function
'There are 9 standard methods for evaluating a quantile in the R language: http://stat.ethz.ch/R-manual/R-patched/library/stats/html/quantile.html
'R's method 7 could be our default, or Langford's methdod 4 'CDF'?
'https://cran.r-project.org/doc/manuals/r-release/fullrefman.pdf
'    Pages 689,1370 and 1533
'more reading:
'http://peltiertech.com/quartiles-for-box-plots/
'http://dsearls.org/other/CalculatingQuartiles/CalculatingQuartiles.htm
'http://superuser.com/questions/343339/excel-quartile-function-doesnt-work
    'try -> Quartile(Split("1,2,3,4",","),QuartileFirst,QmExclusive)
    If Not fArrayIsPresorted Then
        SortArrayInPlace ary
    End If
    Select Case quart
        Case QuartileMinVal
            Quartile = ary(LBound(ary))
        Case QuartileSecond
            Quartile = MedianOfArray(ary, True)
        Case QuartileFirst, QuartileThird
            Quartile = mGetQuartileFromSortedArray(ary, quart, qMethod)
        Case QuartileMaxVal
            Quartile = ary(UBound(ary))
    End Select
End Function

Private Function mGetQuartileFromSortedArray( _
                    ByRef ary As Variant, _
                    ByRef quart As QuartileType, _
                    ByRef qMethod As QuartileMethod _
)
Dim lngLowerBound As LongPtr: lngLowerBound = LBound(ary)
Dim lngUpperBound As LongPtr: lngUpperBound = UBound(ary)
Dim lngElementCount As LongPtr: lngElementCount = (lngUpperBound - lngLowerBound + 1)
    If lngElementCount Mod 2 = 0 Then 'Sove for even number of elements in the array
        Select Case qMethod
            Case QmExclusive, QmInclusive
                ' methods that use similar methods to solve the Quartile if the data set has an even number of values,
                Select Case quart
                    Case QuartileFirst
                        ary = ArraySubset(ary, lngLowerBound, (lngElementCount / 2) - 1)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                    Case QuartileThird
                        ary = ArraySubset(ary, (lngElementCount / 2) - 1, lngUpperBound)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                End Select
            Case Else
                MsgBox "This method not yet built"
                Exit Function
        End Select
    Else 'Solve for odd number of elements in the array
        Select Case qMethod
            Case QmInclusive
                Select Case quart
                    Case QuartileFirst
                        ary = ArraySubset(ary, lngLowerBound, ((lngElementCount + 1) / 2) - 1)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                    Case QuartileThird
                        ary = ArraySubset(ary, lngUpperBound - (((lngElementCount + 1) / 2) - 1), lngUpperBound)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                End Select
            Case QmExclusive
                Select Case quart
                    Case QuartileFirst
                        ary = ArraySubset(ary, lngLowerBound, ((lngElementCount - 1) / 2) - 1)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                    Case QuartileThird
                        ary = ArraySubset(ary, lngUpperBound - (((lngElementCount - 1) / 2) - 1), lngUpperBound)
                        mGetQuartileFromSortedArray = MedianOfArray(ary, True)
                End Select
            Case TodoQmHoggAndLedolter
            Case TodoQmM_S
            Case TodoQmLohninger
            Case TodoQmJ_F
            Case TodoQmH_L
            Case TodoQmH_L2
            Case TodoQmMinitab
            Case TodoQmSas1
            Case TodoQmSas2
            Case TodoQmSas3
        End Select
    End If
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
    Dim ws As Worksheet
    Selection.Copy
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
    ws.Application.CutCopyMode = False
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

