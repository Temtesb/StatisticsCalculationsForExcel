Attribute VB_Name = "bas_SetEcho"
Option Explicit

' Copyright 2011-2014.  All rights reserved. Bradley M. Gough and Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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
         
Public Function IsCursorWait() As Boolean
    
    If Application.Cursor = xlWait Then
        IsCursorWait = True
    Else
        IsCursorWait = False
    End If
End Function

' Used to disable and enable display updates.  When a chain of procedures is
' called many procedures within the chain may want to disable display updates.
' This procedure is used by each individual function to disable and enable display
' updates.  It ensures that display updates will not be enabled untill the last
' call in the chain occurs.
' fValue: True to enabled and False to disable display updates
' fForce: If True, fValue is applied no matter what and intCount is reset
Public Sub SetEcho(ByRef fValue As Boolean, Optional ByRef fForce As Boolean = False)
On Error Resume Next

Static sintCallCount As Integer

    ' Force passed therefore reset call count so value is applied
    If fForce Then _
        sintCallCount = 0

    If fValue Then
        ' Lower disable display updates count by one
        ' Don't let count get less then zero
        If sintCallCount > 0 Then _
            sintCallCount = sintCallCount - 1
        If sintCallCount = 0 Then
            ' - Enable display updates -
            ' apiLockWindowUpdate causes displayed glitch on Windows XP
            ' this issue is no longer limiting as XP has been retired, may be worth looking into
            'apiLockWindowUpdate 0&
        End If
    Else
        If sintCallCount = 0 Then
            ' - Disable display updates -
            ' Note: We use both methods because testing indicated
            ' different results for each version.  Using both
            ' seems to work all the time.
            ' apiLockWindowUpdate causes displayed glitch on Windows XP
            ' this issue is no longer limiting as XP has been retired, may be worth looking into
            'apiLockWindowUpdate Application.hWndAccessApp
        End If
        ' Increase disable display updates count by one
        sintCallCount = sintCallCount + 1
    End If
End Sub
      
