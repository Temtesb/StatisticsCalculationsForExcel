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
'Unlike GEOMEAN, negative numbers > -1 are included (makes sense for loss rates that are not expected to exceed 100%).
'Still need to confirm that this is the same method used in the /Excel function being used
    Dim a As Variant
    Dim x As Variant
    Dim n As Double
    Dim v As Double
    If IsArray(rng) Then
        a = rng
    Else
        a = Array(rng)
    End If
    For Each x In a 'each element in array
        If mProcessIfIsnumeric(x) Then
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
        Err.Raise 3306, Description:="No non zero values exist in the supplied data set"
    End If
End Function

Public Function HarmonicMean(ary As Variant) 'only accepts positive numbers
'Note: this calculation excludes non-numeric values like Excel's GEOMEAN function.
'This function errors if any value is negative, (using negative values is mathmatically possible, but retuns wildly unexpected results)
'Metrics that are inversely proportional to time should be averaged using the harmonic mean
'Also useful in calculating average resistance of differing resistors (or cpacitance of capacitors) in series
'N/(1/n(1)+1/n(2)...)
Dim varElement As Variant
Dim lngCount As LongPtr
Dim dblDenominator As Double
    For Each varElement In ary
        If mProcessIfIsnumeric(varElement) Then
            If varElement > 0 Then
                lngCount = lngCount + 1
                dblDenominator = dblDenominator + (1 / varElement)
            Else
                If varElement < 0 Then '0's are ignored
                    Err.Raise 3305, Description:="This function does not accept negative values"
                End If
            End If
        End If
    Next
    If lngCount Then
        HarmonicMean = lngCount / (dblDenominator)
    Else
        Err.Raise 3307, Description:="No non zero,postitive values exist in the supplied data set"
    End If
End Function

Private Function mProcessIfIsnumeric(ByRef varElement As Variant) As Boolean ' this also converts text numeric values to cdbl
Dim vt As VbVarType
    If (vt + 1 > 1 And vt <= 7) _
        Or vt = 14 Or vt = 20 _
    Then
        mProcessIfIsnumeric = True
    Else
        If IsNumeric(varElement) Then
            varElement = CDbl(varElement)
            mProcessIfIsnumeric = True
        Else
            mProcessIfIsnumeric = False
        End If
    End If
End Function

Public Function IsRange(obj As Variant) As Boolean 'can be run from excel or access, or on access or excel range objects
    On Error Resume Next
    IsRange = TypeName(obj) = "Range"
End Function

Private Function GetArrayFromUnknownObjectSouce(ByRef obj As Variant, Optional intDimentionRank As Integer = 0, Optional fUseObjectAppViceRunningApp As Boolean = True) As Variant
    'Assign any acceptable obect to our aryToProcess
    Dim aryToProcess As Variant
    On Error Resume Next
    
    If IsObject(obj) Then
        Dim strApplicationName As String
        If fUseObjectAppViceRunningApp Then
            'This method checks the application of the object "Default"
            strApplicationName = obj.Application.Name
        Else
            'This method checks the current application that is running code
            'Using late binding so that we don't have to deal with assigning referances
            Dim oThisApplication As Object: Set oThisApplication = Application
            strApplicationName = oThisApplication.Name
        End If
        Select Case True
            Case InStr(1, strApplicationName, "Excel")
                Select Case True
                    Case IsRange(obj)
                        aryToProcess = ExcelRangeToNumericSafeArray(obj)
                End Select
            Case InStr(1, strApplicationName, "Access")
                Select Case TypeName(obj)
                    Case "Recordset"
                        aryToProcess = obj.GetRows
                Case "TableDef", "QueryDef"
                        aryToProcess = obj.OpenRecordset(obj.Name).GetRows
                End Select
            'Case InStr(1, strApplicationName,  "Word")
            'Case InStr(1, strApplicationName, "Powerpoint")
        End Select
    End If
    If aryToProcess Is Nothing Then
        'if we error anywhere we attempt to set the object directly to an array
        aryToProcess = obj
    End If
    If IsArray(aryToProcess) Then
        If ArrayRank(aryToProcess) > 1 Then
            aryToProcess = ArraySingleFromMultiDimention(aryToProcess, intDimentionRank)
        End If
    Else
        aryToProcess = Array(aryToProcess)
        If Not IsArray(aryToProcess) Then
            Set aryToProcess = Nothing
            'Plan is to accept multidimetional arrays, excel ranges, word tables, access recordsets, tabledefs, and querydef and process dedicated column for quartile
            Err.Raise 3302, Description:="The object passed to the Quartile must be an array, excel range, access recordset, access tabledef, access querydef." 'or word table,"
        End If
    End If
    GetArrayFromUnknownObjectSouce = aryToProcess
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
                    Optional fArrayIsPresorted As Boolean = False, _
                    Optional intDimentionRank As Integer _
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
    'try -> Quartile(Split("1,2,3,4",","),QuartileFirst,QmExclusive) ' this is sorted as text, we expect numeric values
    ary = GetArrayFromUnknownObjectSouce(ary, intDimentionRank)
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
            Case Else
                MsgBox "This method not yet built"
                Exit Function
'           Case TodoQmHoggAndLedolter
'            Case TodoQmM_S
'            Case TodoQmLohninger
'            Case TodoQmJ_F
'            Case TodoQmH_L
'            Case TodoQmH_L2
'            Case TodoQmMinitab
'            Case TodoQmSas1
'            Case TodoQmSas2
'            Case TodoQmSas3
        End Select
    End If
End Function

Private Function PRankXinArray(vaArray As Variant, x As Variant) As Double
'This function only works if we are not required to interpolate the results (x is in the array)
'From http://stackoverflow.com`/questions/4800913/percentrank-algorithm-in-vba
'TODO validate this work!!!
    Dim lLower As Long
    Dim lHigher As Long
    Dim i As Long

    For i = LBound(vaArray, 1) To UBound(vaArray, 1)
        If vaArray(i, 1) < x Then
            lLower = lLower + 1
        ElseIf vaArray(i, 1) > x Then
            lHigher = lHigher + 1
        End If
    Next i

    PRankXinArray = lLower / (lLower + lHigher)

End Function

