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
    
    Public Function MedianOfArray(ByRef ary As Variant, Optional fArrayIsSorted As Boolean = False) As Double
    'Finds the median of any single dimention array
        Dim lngElementCount As LongPtr
        lngElementCount = (UBound(ary) - LBound(ary) + 1)
        If Not fArrayIsSorted Then
            SortArrayInPlace ary
        End If
        If lngElementCount Mod 2 <> 0 Then 'uneven amount of numbers the median is the middle element
            MedianOfArray = ary(UBound(ary) - ((lngElementCount - 1) / 2))
        Else 'even amount of numbers
            Dim lngOne As LongPtr
            Dim lngTwo As LongPtr
            lngOne = ary(UBound(ary) - ((lngElementCount) / 2))
            lngTwo = ary(UBound(ary) - ((lngElementCount) / 2) + 1)
            MedianOfArray = (lngOne + lngTwo) / 2
        End If
    End Function
    
    Public Function MeanFromArray(ByRef ary As Variant) As Double 'single dimention array
        Dim lngElementCount As LongPtr
        lngElementCount = (UBound(ary) - LBound(ary) + 1)
        Dim lngSum As Double
        Dim lngElement
        For lngElement = LBound(ary) To UBound(ary)
            lngSum = lngSum + ary(lngElement)
        Next lngElement
        MeanFromArray = lngSum / lngElementCount
    End Function
    
    Public Function ArraySingleFromMultiDimention(ary As Variant, intDimentionRank As Integer) As Variant
        Dim lngElement As LongPtr
        'Create a single dimentioned array with empty elements (we don't have to redim preserve because we know how many elements we have)
        Dim aryNew() As Variant
        ReDim aryNew(UBound(ary) - LBound(ary) + 1)
        'Fill our new array
        For lngElement = LBound(ary) To UBound(ary)
            aryNew(lngElement) = ary(lngElement, intDimentionRank)
        Next lngElement
        ArraySingleFromMultiDimention = aryNew
    End Function
    
    Public Function ModeSingleFromArray(ByRef ary As Variant) As Variant 'single dimention array
    'If multiple modes exists the first one found will be returned
    'this appears to be the behavior of Excels 'Mode.SNGL' function
    'If there is no mode (max occurance is 1) then returns Null
    'Arrays of any data type that can be sorted are accepted,
    'try -> Debug.print ModeSingleFromArray(Split("1,x,x,3",","))
    Dim varVal As Variant, varValWithMost As Variant
    Dim lngMaxOfCountOfVal As LongPtr, lngCountOfThisVal As LongPtr
    Dim lngElement As LongPtr, lngLowerBoundOfArray As LongPtr, lngUpperBoundOfArray As LongPtr
        SortArrayInPlace ary
        lngLowerBoundOfArray = LBound(ary)
        lngUpperBoundOfArray = UBound(ary)
        For lngElement = lngLowerBoundOfArray To lngUpperBoundOfArray
            varVal = ary(lngElement)
            If lngElement = lngLowerBoundOfArray Then 'first elemetent only
                varValWithMost = varVal
                lngCountOfThisVal = 1
                lngMaxOfCountOfVal = 1
            Else
                If varVal = ary(lngElement - 1) Then
                    lngCountOfThisVal = lngCountOfThisVal + 1
                    If lngCountOfThisVal > lngMaxOfCountOfVal Then
                        varValWithMost = varVal
                        lngMaxOfCountOfVal = lngCountOfThisVal
                    End If
                Else
                    lngCountOfThisVal = 1
                End If
            End If
        Next lngElement
        If lngMaxOfCountOfVal = 1 Then
            ModeSingleFromArray = Null
        Else
            ModeSingleFromArray = varValWithMost
        End If
    End Function
    
    Public Function ModeFromArray(ByRef ary As Variant) As Variant() 'currently ony returns as an array
    'single dimention array
    'returns a single dimention array of all of the modes found if more than one exist
        'Arrays of any data type that can be sorted are accepted,
        'try -> debug.print ModeMultiFromArray(Split("1,x,x,3",","))(1)
    Dim varVal As Variant, varValWithMost As Variant
    Dim lngMaxOfCountOfVal As LongPtr, lngCountOfThisVal As LongPtr
    Dim lngElement As LongPtr, lngLowerBoundOfArray As LongPtr, lngUpperBoundOfArray As LongPtr
    Dim lstMode As New VB_Lib_List
        SortArrayInPlace ary
        lngLowerBoundOfArray = LBound(ary)
        lngUpperBoundOfArray = UBound(ary)
        For lngElement = lngLowerBoundOfArray To lngUpperBoundOfArray
            varVal = ary(lngElement)
            If lngElement = lngLowerBoundOfArray Then 'first elemetent only
                varValWithMost = varVal
                lngCountOfThisVal = 1
                lngMaxOfCountOfVal = 1
            Else
                If varVal = ary(lngElement - 1) Then
                    lngCountOfThisVal = lngCountOfThisVal + 1
                    If lngCountOfThisVal = lngMaxOfCountOfVal Then
                        lstMode.Add varVal ' add to the array of our soulutions
                    End If
                    If lngCountOfThisVal > lngMaxOfCountOfVal Then
                        'wipe our answers array and fill with this value
                        lstMode.Clear
                        lstMode.Add varVal
                        lngMaxOfCountOfVal = lngCountOfThisVal
                    End If
                Else
                    lngCountOfThisVal = 1
                End If
            End If
        Next lngElement
        If lngMaxOfCountOfVal = 1 Then
            ModeFromArray = Null
        Else
            ModeFromArray = lstMode.Items
        End If
    End Function

