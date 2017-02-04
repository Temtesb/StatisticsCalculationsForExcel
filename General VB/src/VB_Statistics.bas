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
           If (vt + 1 > 1 And vt + 1 <= 7) _
               Or vt = 14 Or vt = 20 _
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
               If obj.EOF and obj.BOF Then
                    Err.Raise 3303, Description:="The recordset has no records to process"
                    Exit Function
               Else
                    aryToProcess = obj.GetRows
               End If
          Case TypeName((obj) = "TableDef") Or (TypeName(obj) = "QueryDef")
               aryToProcess = obj.OpenRecordset(obj.Name).GetRows
     End Select
    If IsArray(obj) Then
        If ArrayRank(obj) > 1 Then
            aryToProcess = ArraySingleFromMultiDimention(aryToProcess, intDimentionRank)
        End If
        QuartileFromObject = Quartile(aryToProcess, quart, qMethod)
    Else
        Err.Raise 3302, Description:="The object passed to the Quartile must be an array, excel range, access recordset, access tabledef, access querydef."'or word table,"
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
'From http://stackoverflow.com`/questions/4800913/percentrank-algorithm-in-vba
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

    PRank = lLower / (lLower + lHigher)

End Function

