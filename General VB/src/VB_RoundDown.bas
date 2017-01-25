Attribute VB_Name = "VB_RoundDown"
Option Explicit
' Authored 2014-2016 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
     'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
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

Public Function RoundDown(dblVal As Double, intDigitsAfterDecimal As Integer) As Double
Dim strTempVal As String: strTempVal = CStr(dblVal)
Dim lngDecimalLocation As Long: lngDecimalLocation = InStr(1, strTempVal, ".", vbBinaryCompare)
    If lngDecimalLocation = 0 Then
        RoundDown = dblVal
    Else
        RoundDown = CDbl(Left(strTempVal, InStr(1, strTempVal, ".", vbBinaryCompare) + intDigitsAfterDecimal))
    End If
End Function
