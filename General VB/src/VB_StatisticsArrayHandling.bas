Attribute VB_Name = "VB_StatisticsArrayHandling"
Option Explicit
'Authored 2017 by Jeremy Dean Gerdes
     'Public Domain in the United States of America,
     'any international rights are relinquished under CC0 1.0 <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
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
    
    Public Function MedianFromPresortedArray(ByRef ary As Variant) As Double 'single dimention array
        Dim lngElementCount As Long
        lngElementCount = (UBound(ary) - LBound(ary) + 1)
        If lngElementCount Mod 2 <> 0 Then 'uneven amount of numbers
            MedianFromPresortedArray = ary(UBound(ary) - ((lngElementCount - 1) / 2))
        Else 'even amount of numbers
            Dim lngOne As Long
            Dim lngTwo As Long
            lngOne = ary(UBound(ary) - ((lngElementCount) / 2))
            lngTwo = ary(UBound(ary) - ((lngElementCount) / 2) + 1)
            MedianFromPresortedArray = (lngOne + lngTwo) / 2
        End If
    End Function
    
     Public Function MeanFromArray(ByRef ary As Variant) As Double 'single dimention array
        Dim lngElementCount As Long
        lngElementCount = (UBound(ary) - LBound(ary) + 1)
        Dim lngSum As Double
        Dim lngElement
        For lngElement = LBound(ary) To UBound(ary)
               lngSum = lngSum + ary(lngElement)
        Next lngElement
        MeanFromArray = lngSum / lngElementCount
    End Function
    
    Public Function ArraySingleFromMultiDimention(ary As Variant, intDimentionRank As Integer) As Variant
        Dim lngElement As Long
        'Create a single dimentioned array with empty elements (we don't have to redim preserve because we know how many elements we have)
        Dim aryNew() As Variant
        ReDim aryNew(UBound(ary) - LBound(ary) + 1)
        'Fill our new array
        For lngElement = LBound(ary) To UBound(ary)
            aryNew(lngElement) = ary(lngElement, intDimentionRank)
        Next lngElement
        ArraySingleFromMultiDimention = aryNew
    End Function
    
    
    Public Function ModeSingleFromArray(ByRef ary As Variant, Optional fReturnNoneIfMultiple As Boolean = False) As Double 'single dimention array
    'If multiple modes exists the first one found will be returned,
    'this appears to be the behavior of Excels 'Mode.SNGL' function
    'unless fReturnNoneIfMultiple (not like excel)
        'TODO...
    End Function
    
    Public Function ModeMultiFromArray(ByRef ary As Variant) As Variant 'single dimention array
    'returns an array of all of the modes found if more than one exist
        'TODO...
    End Function
          
          
