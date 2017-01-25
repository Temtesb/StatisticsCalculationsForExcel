Attribute VB_Name = "VB_ToolCode"
Option Explicit

'Authored 2011-2017 by Bradley M. Gough with modifications by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
     'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. ß 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C ß 101
         '...
         'A ìwork of the United States Governmentî is a work prepared by an
         'officer or employee of the United States Government as part of that
         'personís official duties.
         '...

' ********************************************************************************
' Introduction
' ********************************************************************************

' The reference information contained in the declaration section of this module
' is provided so that the procedures within this module may be understood.

' VBA grammar is specified using the Augmented Backus-Naur Form (ABNF) metalanguage
' as defined in [RFC4234].  At least a basic understanding of the ABNF metalanguage
' is required to understand the VBA grammar specification.

' ABNF metalanguage reference:
' http://www.ietf.org/rfc/rfc4234.txt

' VBA grammar specification reference:
' http://msdn.microsoft.com/en-us/library/dd361851(v=PROT.10).aspx

' ********************************************************************************
' ABNF metalanguage Reference
' ********************************************************************************

' Rule Naming:
' The name of a rule is simply the name itself; that is, a sequence of
' characters, beginning with an alphabetic character, and followed by a
' combination of alphabetics, digits, and hyphens (dashes).
' NOTE: Rule names are case-insensitive.
' The names <rulename>, <Rulename>, <RULENAME>, and <rUlENamE> all
' refer to the same rule.
' Unlike original BNF, angle brackets ("<", ">") are not required.

' Rule Form:
' A rule is defined by the following sequence:
' name = elements crlf

' Literals:
' ABNF uses % to indicate number value Hence: CR =  %d13
' ABNF permits the specification of literal text strings directly,
' enclosed in quotation-marks.  Hence: Command = "command string"
' Terminals are specified by one or more numeric characters, with the
' base interpretation of those characters indicated explicitly.  The
' following bases are currently defined:
' b = binary
' d = decimal
' x = hexadecimal'

' Alternatives: Rule1 / Rule2
' Elements separated by a forward slash ("/") are alternatives.

' Sequence Group:(Rule1 Rule2)
' Elements enclosed in parentheses are treated as a single element,
' whose contents are STRICTLY ORDERED.

' Variable Repetition: *Rule
' The operator "*" preceding an element indicates repetition.
' The full form is: <a>*<b>element where <a> and <b> are
' optional decimal values, indicating at least <a> and
' at most <b> occurrences of the element.
' Default values are 0 and infinity so that *<element> allows
' any number, including zero; 1*<element> requires at least one;
' 3*3<element> allows exactly 3 and 1*2<element> allows one or two.

' Specific Repetition: nRule
' A rule of the form: <n>element is equivalent to <n>*<n>element
' That is, exactly <n> occurrences of <element>.  Thus, 2DIGIT is a
' 2-digit number, and 3ALPHA is a string of three alphabetic characters.

' Optional Sequence: [RULE]
' Square brackets enclose an optional element sequence: [foo bar]

' Comment: ; Comment
' A semi-colon starts a comment that continues to the end of line.
' This is a simple way of including useful notes in parallel with the
' specifications.

' ********************************************************************************
' VBA Grammar Reference
' ********************************************************************************

' --------------------------------------------------------------------------------
' White Space Grammar
' --------------------------------------------------------------------------------

' VBA grammar defines white-space as a white-space-character(WSC) or a line-continuation.
' VB.NET grammar does not define line-continuations as white-space but documentation
' states that line-continuations are treated like white-space even though they are not.

' Testing indicates that the following assumptions are valid despite the following
' ABNF white space grammar rules:
' 1. A line-continuation underscore character will have 1 or more leading
'    white-space-characters instead of zero or more.
' 2. DBCS-whitespace and most-Unicode-class-Zs white-space-characters
'    do not occur in code.
' 3. %d8232 and %d8233 line-terminator characters do not occur in code.

' ABNF white space grammar rules:
' white-space(WS) = white-space-character(WSC) / line-continuation
' white-space-character(WSC) = (tab-character / eom-character / space-character / DBCS-whitespace / most-Unicode-class-Zs)
' line-continuation = *white-space-character(WSC) underscore *white-space-character(WSC) line-terminator
' line-terminator = (%d13 %d10) / %d13 / %d10 / %d8232 / %d8233
' tab-character = %d9
' eom-character = %d25
' space-character = %d32
' underscore = %d95
' DBCS-whitespace  = %d8233
' most-Unicode-class-Zs = <all members of Unicode class Zs which are not CP2-characters>

' --------------------------------------------------------------------------------
' Line Grammar
' --------------------------------------------------------------------------------

' A physical line is a single line of code.  A physical line may only be part
' of a statement if line-continuations are used to connect physical lines into
' a single logical line.

' A logical line is a complete statement.  A logical line may be a single
' physical line or multiple physical lines connected using line-continuations.

' Testing indicates that the following assumptions are valid despite the following
' ABNF line grammar rules:
' 1. Physical lines will always end with a (%d13 %d10) line-terminator
'    that is not preceded by a line-continuation underscore character.
' 2. A line-continuation underscore character will have 1 or more leading
'    white-space-characters instead of zero or more.
' 3. There will be 0 white-space characters after a line-continuation
'    underscore character instead of one or more.
' 4. There will always be a (%d13 %d10) line-terminator following a
'    line-continuation underscore character.
' 5. DBCS-whitespace and most-Unicode-class-Zs white-space-characters
'    do not occur in code.
' 6. %d8232 and %d8233 line-terminator characters do not occur in code.

' ABNF physical line grammar rules:

' module-body-physical-structure = *source-line [non-terminated-line]
' source-line = *non-line-termination-character line-terminator
' non-terminated-line = *non-line-termination-character
' line-terminator = (%d13 %d10) / %d13 / %d10 / %d8232 / %d8233
' non-line-termination-character =  <any character other than %d13 / %d10 / %d8232 / %d8233>

' ABNF logical line grammar rules:

' module-body-logical-structure = *extended-line
' extended-line =  *(line-continuation / non-line-termination-character) line-terminator
' line-continuation = *white-space-character(WSC) underscore *white-space-character(WSC) line-terminator

' white-space(WS) = white-space-character(WSC) / line-continuation
' white-space-character(WSC) = (tab-character / eom-character / space-character / DBCS-whitespace / most-Unicode-class-Zs)

' tab-character = %d9
' eom-character = %d25
' space-character = %d32
' underscore = %d95
' DBCS-whitespace  = %d8233
' most-Unicode-class-Zs = <all members of Unicode class Zs which are not CP2-characters>

' --------------------------------------------------------------------------------
' Token Grammar
' --------------------------------------------------------------------------------

' Tokens:
' A token is an identifier, keyword, literal, separator, or operator

' Special Tokens:
' Special tokens are used to identify single characters that have special meaing
' in the syntax of a VBA program.  They may be preceded by white-space-characters
' that are ignored.
' Separator special token list:              = '(', ')', ',', '.', '!', '#', ':', '?'
' Arithmetic Operator special token list:    = '^', '*', '/', '\', '+', '-'
' Comparison Operator special token list:    = '<', '=', '>'
' Concatenation Operator special token list: = '&', '+'
' End of statement special token list:       = ';'
' Start of comment special token list:       = ''' (i.e. single-quote)

' Identifier Tokens:
' An identifier is a unique name for an entity such as a variable, procedure,
' class, user-defined data type, or enumeration.
' Identifier naming rules:
' - Must begin with a letter.
' - Must contain only letters, numbers, and underscores (no other symbols).
' - Cannot be more than 255 characters long.
' - Cannot be a keyword word.
' Identifiers must be delimited and may be delimited by any character that is
' not a letter, number, or underscore and has a character code <=127.

' Keyword Tokens:
' A keyword is a reserved word in the VBA that is used to perform a specific task.
' Keyword examples include If, Else, Dim, For, Date, Double, Exit and many others.
' Keywords are special types of identifiers therefore they follow the same naming
' rules and are delimited the same way.

' Literal Tokens:
' A literal is a textual representation of a particular type and value.  Literals
' may represent numbers, dates, or strings.  A literal will be assigned a value at
' compile time, while a variable will be assigned a value at runtime.
' Literals do not need to be delimited if you consider base type prefix and data
' type suffix declaration characters part of number literals, enclosing # characters
' part of date literals and enclosing " characters part of a string literals.
' The following are examples that prove literals do not need to be delimited:
' - Entering 'Debug.Print 2+2' into the debug window will return 4 as expected
'   therefore the '2' number literals do not need to be delimited.
' - Entering 'Debug.Print 5#Mod 3#' into the debug window will return '2' as expected
'   therefore the '5#' number literal did not need to be delimited.
' - Entering 'Debug.Print#1/1/2011#' into the debug window will return '1/1/2011' as expected
'   therefore the '#1/1/2011#' date literal did not need to be delmited.
' - Entering 'Debug.Print"Test"&"Test"' into the debug window will return 'TestTest' as expected
'   therefore the '"Test"' string literals did not need to be delimited.

' Number Literals:
' The following are number literal base prefix characters:
' Hex = &H
' Oct = &O
' The following are number literal data type suffix characters:
' Integer  = %
' Long     = &
' Single   = !
' Double   = #
' Currency = @
' Examples:
' Note: The VBA IDE somtimes automatically deletes a literal
' type declaration character if it is the default type.
' Private Const intInteger = 1%
' Private Const lngLong = 1&
' Private Const sngSingle = 1.1!
' Private Const dblDouble = 1.1#
' Private Const curCurrency = 1@

' Date Literals:
' Date literals begine and end with the # character.
' Examples:
' Private Const dtmDate = #11/11/11#
' Private Const dtmDate = #Jan 01, 2011#

' String Literals:
' String literals begine and end with the " character.
' The "" characters are and escape sequence to represent single " character
' within a string literal.
' Examples:
' Private Const strString = "Test""Test"

' Separator Tokens:
' A separator is a sequence of one or more characters used to specify a boundary.
' Separators do not need to be delimited.
' Separator list: '(', ')', ',', '.', '!', '#', ':', ':=', '?'

' Operator Tokens:
' An operator takes an action on one or more expressions.
' Operators that are special-tokens do not need to be delimited.
' Operators that are not special-tokens such as 'Mod', 'Like', and 'Not'
' must be white-space or special-token delimited.
' The VBA IDE automatically puts white-space around all operators including
' special-token operators but white-space is not required when special-token
' operators are entered into the debug window.
' Arithmetic operator list: '^', '*', '/', '\', 'Mod', '+', '-'
' Comparison operator list: '<', '<=', '>', '>=', '=', '<>', 'Is', 'Like'
' Concatenation operator list: '&', '+'
' Logical operator list: 'And', 'Eqv', 'Imp', 'Not', 'Or', 'Xor'

' --------------------------------------------------------------------------------
' Comment Grammar
' --------------------------------------------------------------------------------

' Comments are used by humans to clarify the source code.
' All comments in VBA follow either the single-quote character or the Rem keyword.

' The first single-quote character within a logical line that is not inside a
' literal (string or date) marks the start position of a comment that will
' continue until the end of the logical line.

' ********************************************************************************
' VBA Statement Reference
' ********************************************************************************

' --------------------------------------------------------------------------------
' Declarations section of code module reference
' --------------------------------------------------------------------------------

' Variable declaration syntaxes:
' Public [WithEvents] varname[([subscripts])] [As [New] type] [,[WithEvents] varname[([subscripts])] [As [New] type]] . . .
' Private [WithEvents] varname[([subscripts])] [As [New] type] [,[WithEvents] varname[([subscripts])] [As [New] type]] . . .
' Dim [WithEvents] varname[([subscripts])] [As [New] type] [, [WithEvents] varname[([subscripts])] [As [New] type]] .
' [Public | Private] Const constname [As type] = expression

' User-defined data type declaration syntax:
' [Private | Public] Type varname
'   elementname [([subscripts])] As type
'   [elementname [([subscripts])] As type]
'   . . .
' End Type

' Enumeration declaration syntax:
' [Public | Private] Enum name
'   membername [= constantexpression]
'   membername [= constantexpression]
'   . . .
' End Enum

' User-defined event declaration syntax:
' [Public] Event procedurename [(arglist)]

' Reference to external procedures in a dynamic-link library (DLL) declaration syntaxes:
' [Public | Private] Declare Sub name Lib "libname" [Alias "aliasname"] [([arglist])]
' [Public | Private] Declare Function name Lib "libname" [Alias "aliasname"] [([arglist])] [As type]

' --------------------------------------------------------------------------------
' Procedure section of code module reference
' --------------------------------------------------------------------------------

' Procedure declaration syntaxes syntaxes:
' [Public | Private | Friend] [Static] Sub name [(arglist)]
' [Public | Private | Friend] [Static] Function name [(arglist)] [As type]
' [Public | Private | Friend] [Static] Property Get name [(arglist)] [As type]
' [Public | Private | Friend] [Static] Property Let name ([arglist,] value)
' [Public | Private | Friend] [Static] Property Set name ([arglist,] reference)

' Procedure variable declaration syntaxes:
' Dim [WithEvents] varname[([subscripts])] [As [New] type] [, [WithEvents] varname[([subscripts])] [As [New] type]] .
' Static varname[([subscripts])] [As [New] type] [, varname[([subscripts])] [As [New] type]] . . .
' Const constname [As type] = expression

Public Enum tttcToolTokenType
    tttcIdentifier = 1
    tttcKeyword = 2
    tttcNumberLiteral = 3
    tttcDateLiteral = 4
    tttcStringLiteral = 5
    tttcSeparator = 6
    tttcOperator = 7
End Enum

Public Sub ToolFindCodeModulesWithCode(ByRef strFind As String)

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
            Dim lngStartLine As Long
            Dim lngStartColumn As Long
            lngStartLine = 1
            lngStartColumn = 1
            Do While .Find(strFind, lngStartLine, lngStartColumn, -1, -1, True, True, False)
                .CodePane.Show
                Debug.Print objVBComponent.Name & " line = " & lngStartLine
                lngStartColumn = lngStartColumn + 1
            Loop
        End With
    Next

End Sub

Public Sub ToolFindCodeModulesWithoutCode(ByRef strFind As String)

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
            Dim fFound As Boolean
            fFound = False
            Do While .Find(strFind, 1, 1, -1, -1, True, True, False)
                fFound = True
                Exit Do
            Loop
            If Not fFound Then
                .CodePane.Show
                Debug.Print objVBComponent.Name

            End If
        End With
    Next

End Sub

'Find lines of code containing strXCode but missing strYCode
'ToolFindCodeModulesWithXCodeAndWithoutYCode "StrComp", "vbBinaryCompare"
Public Sub ToolFindCodeModulesWithXCodeAndWithoutYCode(ByRef strXCode As String, ByRef strYCode As String)
    Dim strLog As String
    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
          Dim lngStartLine As Long
            Dim lngStartColumn As Long
            lngStartLine = 1
            lngStartColumn = 1
            Do While .Find(strXCode, lngStartLine, lngStartColumn, -1, -1, True, True, False)
                Dim strFoundLogicalLine As String
                Dim lngMultiLineStatementLineCount As Long
                strFoundLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngStartLine, lngMultiLineStatementLineCount)
                If InStr(1, strFoundLogicalLine, strYCode, vbBinaryCompare) = 0 Then
                    .CodePane.Show
                    Debug.Print objVBComponent.Name & " line = " & lngStartLine & " :containing " & _
                        strXCode & " without " & strYCode & vbCrLf & vbTab & strFoundLogicalLine
                    strLog = strLog & vbCrLf & objVBComponent.Name & " line = " & lngStartLine & " :containing " & _
                        strXCode & " without " & strYCode & vbCrLf & vbTab & strFoundLogicalLine
                End If
            lngStartColumn = lngStartColumn + 1
            Loop
        End With
    Next
    
    ToolSaveStringToFile Environ$("TEMP") & "\Desktop\Log.txt", strLog

End Sub

Private Sub ToolSaveStringToFile(strPath As String, strMsg As String)
Dim objFile As Object
Dim fso As Object: fso = CreateObject("Scripting.FileSystemObject")
    Set objFile = fso.CreateTextFile(strPath, True)
    objFile.Write strMsg
    objFile.Close
    On Error Resume Next
    Shell "explorer.exe " & strPath, vbNormalFocus
End Sub

Public Sub ToolRemoveSpacesFromBlankLinesOfCode()

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
            Dim lngLine As Long
            For lngLine = 1 To .CountOfLines
                Dim strLine As String
                strLine = .Lines(lngLine, 1)
                If LenB(strLine) <> 0 Then
                    strLine = Trim$(.Lines(lngLine, 1))
                    If LenB(strLine) = 0 Then
                        Debug.Print objVBComponent.Name & "; Spaces removed from line " & lngLine
                        ' Note: Cannot use vbNullstring to clear line, must use ""
                        .ReplaceLine lngLine, ""
                    End If
                End If
            Next
        End With
    Next

End Sub

Public Sub ToolFindMultiLineStatements()

    Dim strLog As String

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
            Dim lngLine As Long
            For lngLine = 1 To .CountOfLines
                ' Test for multi-line statement by looking for line continuation character
                Dim strLine As String
                strLine = .Lines(lngLine, 1)
                If Right$(strLine, 2) = " _" Then

                    Dim strMultiLineStatement As String
                    Dim lngMultiLineStatementLineCount As Long
                    strMultiLineStatement = ToolGetLogicalLine(objVBComponent.CodeModule, lngLine, lngMultiLineStatementLineCount)

                    #If True Then
                        Debug.Print objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf
                        strLog = strLog & objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf & vbCrLf
                    #ElseIf False Then
                        ' Filter out statements that probably have line breaks at the correct locations
                        If Right$(strLine, 7) <> " Then _" And _
                            Right$(strLine, 3) <> "( _" And _
                            Right$(strLine, 4) <> " = _" And _
                            Right$(strLine, 4) <> " & _" And _
                            Right$(strLine, 5) <> " Or _" And _
                            Right$(strLine, 6) <> " And _" Then
                            Debug.Print objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf
                            strLog = strLog & objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf & vbCrLf
                        End If
                    #Else
                        ' Find statements that operators at the start of a continuation line instead
                        ' of the end of the line being continued
                        Dim lngMultlineStatementLine As Long
                        For lngMultlineStatementLine = lngLine To lngLine + lngMultiLineStatementLineCount - 1
                            Dim strMultiLineStatementLine As String
                            strMultiLineStatementLine = Trim$(.Lines(lngMultlineStatementLine, 1))
                            If Left$(strMultiLineStatementLine, 4) = "And " Or _
                                Left$(strMultiLineStatementLine, 3) = "Or " Or _
                                Left$(strMultiLineStatementLine, 4) = "Xor " Or _
                                Left$(strMultiLineStatementLine, 4) = "Eqv " Or _
                                Left$(strMultiLineStatementLine, 4) = "Imp " Or _
                                Left$(strMultiLineStatementLine, 2) = "& " Or _
                                Left$(strMultiLineStatementLine, 2) = "< " Or _
                                Left$(strMultiLineStatementLine, 2) = "> " Or _
                                Left$(strMultiLineStatementLine, 2) = "= " Or _
                                Left$(strMultiLineStatementLine, 2) = "+ " Or _
                                Left$(strMultiLineStatementLine, 2) = "- " Or _
                                Left$(strMultiLineStatementLine, 2) = "* " Or _
                                Left$(strMultiLineStatementLine, 2) = "/ " Or _
                                Left$(strMultiLineStatementLine, 2) = "\ " Or _
                                Left$(strMultiLineStatementLine, 2) = "^ " Or _
                                Left$(strMultiLineStatementLine, 4) = "Mod " Then
                                Debug.Print objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf
                                strLog = strLog & objVBComponent.Name & vbCrLf & strMultiLineStatement & vbCrLf & vbCrLf
                            End If
                        Next
                    #End If

                    ' Increment line to last line of multi-line statement
                    lngLine = lngLine + lngMultiLineStatementLineCount - 1

                End If
            Next
        End With
    Next

    ToolSaveStringToFile Environ$("TEMP") & "\Desktop\Log.txt", strLog

End Sub



Public Sub ToolSetIndentForMultiLineStatements()

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule
            Dim lngLine As Long
            For lngLine = 1 To .CountOfLines
                Dim strLine As String
                strLine = .Lines(lngLine, 1)
                ' Test for multi-line statement by looking for line continuation character
                If Right$(strLine, 2) = " _" Then
                    ' Get indent position for all lines after first line of multi-line statement
                    Dim lngIndentPosition As Long
                    lngIndentPosition = ToolGetFirstTokenPosition(strLine, 1, True)
                    lngIndentPosition = lngIndentPosition + 4
                    'Debug.Print strLine
                    Do
                        ' Move to next line
                        lngLine = lngLine + 1
                        ' Get indented line
                        strLine = Space(lngIndentPosition - 1) & Trim$(.Lines(lngLine, 1))
                        ' Replace line with newly indented line
                        .ReplaceLine lngLine, strLine
                        'Debug.Print strLine
                        ' If line does not contain a line continuation character then we have
                        ' reached the end of the statement and need to stop indenting lines
                        If Right$(strLine, 2) <> " _" Then
                            Exit Do
                        End If
                    Loop
                End If
            Next
        End With
    Next

End Sub

Public Sub ToolRemoveCallKeywordFromProcedureCalls()

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents

        With objVBComponent.CodeModule

            Dim lngStartLine As Long
            Dim lngStartColumn As Long
            lngStartLine = 1
            lngStartColumn = 1

            Do While .Find("Call", lngStartLine, lngStartColumn, -1, -1, True, True, False)

                ' Get logical line
                Dim strLogicalLine As String
                strLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngStartLine)

                Select Case VBA.MsgBox("Remove Call?" & vbCrLf & vbCrLf & objVBComponent.Name & vbCrLf & vbCrLf & strLogicalLine, vbQuestion + vbYesNoCancel)

                    Case vbYes

                        ' Overwrite 'Call' keywork with 4 spaces
                        Mid$(strLogicalLine, lngStartColumn, 4) = "    "
                        ' Overwrite start argument '(' with 1 space
                        Dim lngStartArgumentParathesisPosition As Long
                        lngStartArgumentParathesisPosition = InStr(lngStartColumn + 5, strLogicalLine, "(", vbBinaryCompare)
                        Mid$(strLogicalLine, lngStartArgumentParathesisPosition, 1) = " "
                        ' Remove end argument ')'
                        Dim lngEndArgumentParathesisPosition As Long
                        lngEndArgumentParathesisPosition = InStrRev(strLogicalLine, ")", -1, vbBinaryCompare)
                        strLogicalLine = Left$(strLogicalLine, lngEndArgumentParathesisPosition - 1) & Right$(strLogicalLine, Len(strLogicalLine) - lngEndArgumentParathesisPosition)

                        ' Split logical line into physical lines
                        Dim astrPhysicalLines() As String
                        astrPhysicalLines = Split(strLogicalLine, vbCrLf, -1, vbBinaryCompare)

                        ' Shift all physical lines to the left 5 spaces to account for removal of 'Call '
                        Dim lngPhysicalLineIndex As Long
                        For lngPhysicalLineIndex = 0 To UBound(astrPhysicalLines)
                            Dim strPhysicalLine As String
                            strPhysicalLine = astrPhysicalLines(lngPhysicalLineIndex)
                            astrPhysicalLines(lngPhysicalLineIndex) = Right$(strPhysicalLine, Len(strPhysicalLine) - 5)
                        Next

                        ' Join physical lines into logical line
                        strLogicalLine = Join(astrPhysicalLines, vbCrLf)

                        ' Replace current logical line with new logical line (with Call keyword removed)
                        ToolDeleteLogicalLine objVBComponent.CodeModule, lngStartLine
                        .InsertLines lngStartLine, strLogicalLine
                        Debug.Print "Call Removed: " & objVBComponent.Name & vbCrLf & strLogicalLine

                    Case vbNo
                        Debug.Print "Call Not Removed: " & objVBComponent.Name & vbCrLf & strLogicalLine

                    Case vbCancel
                        Exit Sub

                End Select

                ' Increment line and reset column so we can search for next call statement
                lngStartLine = lngStartLine + 1
                lngStartColumn = 1

            Loop

        End With

    Next

End Sub

Public Sub ToolAddByRefKeywordToProcedureDeclarations()

    Dim strLog As String

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule

            ' Cycle through all non declaration lines of code
            Dim lngNonDeclarationLine As Long
            For lngNonDeclarationLine = .CountOfDeclarationLines + 1 To .CountOfLines

                ' Get procedure name and type for current line of code
                Dim strProcedureName As String
                Dim lngProcedureType As Long ' vbext_ProcKind
                strProcedureName = .ProcOfLine(lngNonDeclarationLine, lngProcedureType)

                ' Verify we found a procedure name
                If LenB(strProcedureName) <> 0 Then

                    'Debug.Print strProcedureName

                    ' Verify procedure is not an event procedure
                    ' Note: This procedure does not modify event procedures
                    ' because event procedure declarations are automatically
                    ' created therefore this procedure leaves them alone.
                    ' Note: We assume all procedures with '_' in their name are
                    ' event procedures.  All event procedures will have '_' in
                    ' their name but some procedures may have '_' in their name
                    ' without being an event procedure.
                    If InStr(1, strProcedureName, "_", vbBinaryCompare) = 0 Then

                        'Debug.Print strProcedureName

                        ' Reset replace procedure declaration statement flag
                        Dim fReplaceProcedureDeclarationStatement As Boolean
                        fReplaceProcedureDeclarationStatement = False

                        ' Get procedure declaration start line
                        Dim lngProcedureDeclarationStatementStartLine As Long
                        lngProcedureDeclarationStatementStartLine = .ProcBodyLine(strProcedureName, lngProcedureType)

                        ' Get procedure declaration statement
                        Dim strProcedureDeclarationStatement As String
                        strProcedureDeclarationStatement = ToolGetLogicalLine(objVBComponent.CodeModule, lngProcedureDeclarationStatementStartLine)
                        'Debug.Print objVBComponent.Name & "; Line " & lngProcedureDeclarationStatementStartLine & "; " & strProcedureDeclarationStatement

                        ' Find delimiter that indicates start of arguments
                        ' Note: Find '(' delimiter
                        Dim lngProcedureDeclarationStatementPosition As Long
                        lngProcedureDeclarationStatementPosition = ToolGetTokenPosition( _
                            ToolGetTokenSearchLogicalLine(strProcedureDeclarationStatement), _
                            1, True, "(", tttcSeparator)

                        ' Set position to 1 character after delimiter that indicates start of arguments
                        ' Note: Set position to 1 character after '(' delimiter
                        lngProcedureDeclarationStatementPosition = lngProcedureDeclarationStatementPosition + 1

                        ' Find position of next token after '(' delimiter
                        lngProcedureDeclarationStatementPosition = ToolGetFirstTokenPosition( _
                            strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, True)

                        ' Verify position is not at the delimiter that indicates the end of arguments
                        ' Note: Position not at ')' delimiter
                        If Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 1) <> ")" Then

                            ' Set position to first character after "Optional " if needed
                            ' Note: Code does not support line continuation character
                            ' between Optional keyword and remainder of argument.
                            If Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 9) = "Optional " Then
                                lngProcedureDeclarationStatementPosition = lngProcedureDeclarationStatementPosition + 9
                            End If

                            ' Insert ByRef if argument does not already start with ByRef, ByVal or ParamArray
                            If Not (Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 6) = "ByRef " Or _
                                Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 6) = "ByVal " Or _
                                Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 11) = "ParamArray ") Then
                                ' Insert ByRef
                                fReplaceProcedureDeclarationStatement = True
                                'Debug.Print strProcedureDeclarationStatement
                                strProcedureDeclarationStatement = _
                                    Left$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition - 1) & _
                                    "ByRef " & _
                                    Right$(strProcedureDeclarationStatement, Len(strProcedureDeclarationStatement) - lngProcedureDeclarationStatementPosition + 1)
                                'Debug.Print strProcedureDeclarationStatement
                            End If

                            Do

                                ' Note: The following code does not support commas inside
                                ' comments trailing a procedure declaration.
                                ' i.e. Public Sub ProcureName(Arg1 As String) ' , Comment

                                ' Find delimiter that indicates start of next argument
                                ' Note: Find ',' delimiter
                                lngProcedureDeclarationStatementPosition = ToolGetTokenPosition( _
                                    ToolGetTokenSearchLogicalLine(strProcedureDeclarationStatement), _
                                    lngProcedureDeclarationStatementPosition, True, ",", tttcSeparator)

                                ' Verify delimiter that indicates start of next argument found
                                If lngProcedureDeclarationStatementPosition <> 0 Then

                                    ' Set position to 1 character after delimiter that indicates start of next argument
                                    ' Note: Set position to 1 character after ',' delimiter
                                    lngProcedureDeclarationStatementPosition = lngProcedureDeclarationStatementPosition + 1

                                    ' Find position of next token after ',' delimiter
                                    lngProcedureDeclarationStatementPosition = ToolGetFirstTokenPosition( _
                                        strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, True)

                                    ' Set position to first character after "Optional " if needed
                                    ' Note: Code does not support line continuation character
                                    ' between Optional keyword and remainder of argument.
                                    If Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 9) = "Optional " Then
                                        lngProcedureDeclarationStatementPosition = lngProcedureDeclarationStatementPosition + 9
                                    End If

                                    ' Insert ByRef if argument does not already start with ByRef, ByVal or ParamArray
                                    If Not (Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 6) = "ByRef " Or _
                                        Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 6) = "ByVal " Or _
                                        Mid$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition, 11) = "ParamArray ") Then
                                        ' Insert ByRef
                                        fReplaceProcedureDeclarationStatement = True
                                        'Debug.Print strProcedureDeclarationStatement
                                        strProcedureDeclarationStatement = _
                                            Left$(strProcedureDeclarationStatement, lngProcedureDeclarationStatementPosition - 1) & _
                                            "ByRef " & _
                                            Right$(strProcedureDeclarationStatement, Len(strProcedureDeclarationStatement) - lngProcedureDeclarationStatementPosition + 1)
                                        'Debug.Print strProcedureDeclarationStatement
                                    End If

                                Else
                                    ' Failed to find anymore arguments therefore exit loop
                                    Exit Do
                                End If

                            Loop

                        End If

                        If fReplaceProcedureDeclarationStatement Then

                            Debug.Print objVBComponent.Name & vbCrLf & strProcedureDeclarationStatement & vbCrLf
                            strLog = strLog & objVBComponent.Name & vbCrLf & strProcedureDeclarationStatement & vbCrLf & vbCrLf

                            ' Replace current procedure declaration statement with new statement
                            ToolDeleteLogicalLine objVBComponent.CodeModule, lngProcedureDeclarationStatementStartLine
                            .InsertLines lngProcedureDeclarationStatementStartLine, strProcedureDeclarationStatement

                        End If

                    End If

                    ' Skip remaining lines of code in procedure
                    lngNonDeclarationLine = lngNonDeclarationLine + .ProcCountLines(strProcedureName, lngProcedureType) - 1

                End If

            Next

        End With
    Next

    ToolSaveStringToFile Environ$("TEMP") & "\Desktop\Log.txt", strLog

End Sub

Public Sub ToolFindDeclarationsWithoutDefinedScope()

    ' Purpose:
    ' Find declarations without defined scope so that they may be fixed.

    ' Restrictions:
    ' Procedure may incorrectly find declarations without defined scope.
    ' Specifically, user defined data types may have members that are keywords
    ' such as 'Const', 'Dim', and 'Event'.  This is rare therefore this procedure
    ' is good enough to get the job done.

    Dim strLog As String

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule

            ' Cycle through all declaration lines of code
            Dim lngDeclarationLine As Long
            For lngDeclarationLine = 1 To .CountOfDeclarationLines

                ' Test for keywords that indicate declaration statement without defined scope
                Dim strDeclarationPhysicalLine As String
                strDeclarationPhysicalLine = LTrim$(.Lines(lngDeclarationLine, 1))
                If Left$(strDeclarationPhysicalLine, 6) = "Const " Or _
                    Left$(strDeclarationPhysicalLine, 4) = "Dim " Or _
                    Left$(strDeclarationPhysicalLine, 5) = "Type " Or _
                    Left$(strDeclarationPhysicalLine, 5) = "Enum " Or _
                    Left$(strDeclarationPhysicalLine, 6) = "Event " Or _
                    Left$(strDeclarationPhysicalLine, 8) = "Declare " Then

                    Dim strDeclarationLogicalLine As String
                    strDeclarationLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngDeclarationLine)

                    Debug.Print objVBComponent.Name & vbCrLf & strDeclarationLogicalLine & vbCrLf
                    strLog = strLog & objVBComponent.Name & vbCrLf & strDeclarationLogicalLine & vbCrLf & vbCrLf

                End If

            Next

            ' Cycle through all non declaration lines of code
            Dim lngNonDeclarationLine As Long
            For lngNonDeclarationLine = .CountOfDeclarationLines + 1 To .CountOfLines

                ' Get procedure name for current line of code
                Dim strProcedureName As String
                Dim lngProcedureType As Long ' vbext_ProcKind
                strProcedureName = .ProcOfLine(lngNonDeclarationLine, lngProcedureType)

                ' Verify we found a procedure name
                If LenB(strProcedureName) <> 0 Then

                    ' Get procedure start line
                    Dim lngProcedureStartLine As Long
                    lngProcedureStartLine = .ProcBodyLine(strProcedureName, lngProcedureType)

                    ' Test for keywords that indicate procedure declaration without defined scope
                    Dim strProcedurePhysicalLine As String
                    strProcedurePhysicalLine = LTrim$(.Lines(lngProcedureStartLine, 1))
                    If Left$(strProcedurePhysicalLine, 7) = "Static " Or _
                        Left$(strProcedurePhysicalLine, 4) = "Sub " Or _
                        Left$(strProcedurePhysicalLine, 9) = "Function " Or _
                        Left$(strProcedurePhysicalLine, 9) = "Property " Then

                        Dim strProcedureLogicalLine As String
                        strProcedureLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngProcedureStartLine)

                        Debug.Print objVBComponent.Name & vbCrLf & strProcedureLogicalLine & vbCrLf
                        strLog = strLog & objVBComponent.Name & vbCrLf & strProcedureLogicalLine & vbCrLf & vbCrLf

                    End If

                    ' Skip remaining lines of code in procedure
                    lngNonDeclarationLine = lngNonDeclarationLine + .ProcCountLines(strProcedureName, lngProcedureType) - 1

                End If

            Next

        End With
    Next

    ToolSaveStringToFile Environ$("TEMP") & "\Desktop\Log.txt", strLog

End Sub

Public Sub ToolFindDeclarationsWithoutDefinedDataType()

    ' Purpose:
    ' Find declarations without defined data type so that they may be fixed.

    ' Restrictions:
    ' Procedure may incorrectly find user defined type declarations without
    ' defined scope if user defined type declaration contains one or more
    ' entire line comments.  This is rare therefore this procedure is good
    ' enough to get the job done.

    ' Assumptions:
    ' All physical line terminators will be vbCrLf when working within the VBA IDE.

    Dim strLog As String

    Dim objVBComponent As Object ' VBIDE.VBComponent
    For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
        With objVBComponent.CodeModule

            Dim fDeclarationWithoutDefinedDataType As Boolean

            Dim lngLine As Long
            Dim strLogicalLine As String
            Dim lngLogicalLinePhysicalLineCount As Long
            Dim strSearchLogicalLine As String

            Dim lngCompoundStatementLine As Long
            Dim StrCompoundStatementLogicalLine As String
            Dim lngCompoundStatementLogicalLinePhysicalLineCount As Long
            Dim StrCompoundStatementSearchLogicalLine As String
            Dim lngCompoundStatementPhysicalLineCount As Long
            Dim StrCompoundStatementPhysicalLines As String

            Dim strProcedureName As String
            Dim lngProcedureType As Long ' vbext_ProcKind
            Dim lngProcedureStartLine As Long
            Dim lngProcedureLineCount As Long
            Dim lngProcedureLine As Long

            Dim lngArgumentsStartPosition As Long
            Dim lngArgumentsEndPosition As Long
            Dim strSearchArguments As String
            Dim astrSearchArguments() As String
            Dim lngSearchArgumentIndex As Long
            Dim lngReturnDataTypeStartPosition As Long

            Dim strSearchTypeElements As String
            Dim astrSearchTypeElements() As String
            Dim lngSearchTypeElementIndex As Long

            Dim astrSearchVariables() As String
            Dim lngSearchVariableIndex As Long

            ' Cycle through all declaration lines of code
            For lngLine = 1 To .CountOfDeclarationLines

                ' Reset declaration without defined data type flag
                fDeclarationWithoutDefinedDataType = False

                ' Get logical line and count of physical lines in logical line
                strLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngLine, lngLogicalLinePhysicalLineCount)

                ' Get search logical line for use when searching for token positions
                strSearchLogicalLine = ToolGetTokenSearchLogicalLine(strLogicalLine)

                ' Reset compound statement physical line count
                lngCompoundStatementPhysicalLineCount = 0

                ' Test if logical line is:
                ' 1. user-defined event procedure declaration
                ' 2. reference to external procedure in DLL declaration or
                If ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Event", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Declare", tttcKeyword) <> 0 Then

                    ' Get arguments start and end positions
                    lngArgumentsStartPosition = ToolGetTokenPosition(strSearchLogicalLine, 1, True, "(", tttcSeparator)
                    lngArgumentsStartPosition = lngArgumentsStartPosition + 1
                    lngArgumentsEndPosition = ToolGetTokenPosition(strSearchLogicalLine, -1, False, ")", tttcSeparator)
                    lngArgumentsEndPosition = lngArgumentsEndPosition - 1
                    ' Verify one or more arguments exists or procedure returns and array
                    If lngArgumentsEndPosition > lngArgumentsStartPosition Then
                        ' Test if procedure returns and array
                        If Mid$(strSearchLogicalLine, lngArgumentsEndPosition, 1) = "(" Then
                            ' Procedure returns array therefore need to find next ')'
                            ' searching in reverse to find arguments end position
                            lngArgumentsEndPosition = ToolGetTokenPosition(strSearchLogicalLine, lngArgumentsEndPosition, False, ")", tttcSeparator)
                            lngArgumentsEndPosition = lngArgumentsEndPosition - 1
                        End If
                    End If
                    ' Verify one or more arguments exists
                    If lngArgumentsEndPosition > lngArgumentsStartPosition Then
                        ' Get search arguments string
                        strSearchArguments = Mid$(strSearchLogicalLine, lngArgumentsStartPosition, lngArgumentsEndPosition - lngArgumentsStartPosition + 1)
                        ' Split search arguments string into search arguments array containing one search argument per element
                        astrSearchArguments() = Split(strSearchArguments, ",", -1, vbBinaryCompare)
                        ' Cycle through all search arguments in array
                        For lngSearchArgumentIndex = 0 To UBound(astrSearchArguments)
                            ' Test if argument data type is missing
                            If ToolGetTokenPosition(astrSearchArguments(lngSearchArgumentIndex), 1, True, "As", tttcKeyword) = 0 Then
                                fDeclarationWithoutDefinedDataType = True
                                Exit For
                            End If
                        Next
                        ' Verify we have not already determined that data type is missing for at least one argument
                        If Not fDeclarationWithoutDefinedDataType Then
                            ' Test if declaration is for procedure with return data type
                            If ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Function", tttcKeyword) <> 0 Then
                                ' Get return data type start position
                                lngReturnDataTypeStartPosition = lngArgumentsEndPosition + 2
                                ' Test if return data type is missing
                                If ToolGetTokenPosition(strSearchLogicalLine, lngReturnDataTypeStartPosition, True, "As", tttcKeyword) = 0 Then
                                    fDeclarationWithoutDefinedDataType = True
                                End If
                            End If
                        End If
                    End If

                ' Test if logical line is user-defined data type declaration
                ElseIf ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Type", tttcKeyword) <> 0 Then

                    ' Get compound statement physical line count
                    For lngCompoundStatementLine = lngLine + lngLogicalLinePhysicalLineCount To .CountOfDeclarationLines
                        ' Get compound statement logical line and count of physical lines in logical line
                        StrCompoundStatementLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngCompoundStatementLine, lngCompoundStatementLogicalLinePhysicalLineCount)
                        ' Get compound statement search logical line for use when searching for token positions
                        StrCompoundStatementSearchLogicalLine = ToolGetTokenSearchLogicalLine(StrCompoundStatementLogicalLine)
                        ' Test if at end of compound statement
                        ' Note: Code does not support line continuation between End and Type keywords.
                        If ToolGetTokenPosition(StrCompoundStatementSearchLogicalLine, 1, True, "End Type", tttcKeyword) <> 0 Then
                            lngCompoundStatementPhysicalLineCount = lngCompoundStatementLine - lngLine + 1
                            Exit For
                        End If
                        ' Set current compound statement line to last physical line of logical line so next
                        ' loop iteration will check next physical line after compound statement logical line
                        lngCompoundStatementLine = lngCompoundStatementLine + lngCompoundStatementLogicalLinePhysicalLineCount - 1
                    Next

                    ' Verify we found compound statement physical line count
                    If lngCompoundStatementPhysicalLineCount <> 0 Then

                        ' Get compound statement physical lines string
                        StrCompoundStatementPhysicalLines = objVBComponent.CodeModule.Lines(lngLine, lngCompoundStatementPhysicalLineCount)

                        ' Get search type elements string
                        strSearchTypeElements = ToolGetTokenSearchLogicalLine(objVBComponent.CodeModule.Lines(lngLine + lngLogicalLinePhysicalLineCount, lngCompoundStatementPhysicalLineCount - lngLogicalLinePhysicalLineCount - lngCompoundStatementLogicalLinePhysicalLineCount))

                        ' Convert comma delimiters to physical line terminator delimiters
                        ' Note: Type elements may be separated by comma delimiters or
                        ' physical line terminator therefore we convert all comma delimiters
                        ' tophysical line terminators to make it easy to split type elements
                        ' into an array.
                        strSearchTypeElements = Replace(strSearchTypeElements, ",", vbCrLf, 1, -1, vbBinaryCompare)
                        ' Split search type elements string into search type elements array
                        ' containing one search type element per array element
                        astrSearchTypeElements() = Split(strSearchTypeElements, vbCrLf, -1, vbBinaryCompare)
                        ' Cycle through all search type elements in array
                        For lngSearchTypeElementIndex = 0 To UBound(astrSearchTypeElements)
                            ' Test if argument data type is missing
                            If ToolGetTokenPosition(astrSearchTypeElements(lngSearchTypeElementIndex), 1, True, "As", tttcKeyword) = 0 Then
                                fDeclarationWithoutDefinedDataType = True
                                Exit For
                            End If
                        Next
                    End If

                ' Test if logical line is enumeration declaration
                ElseIf ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Enum", tttcKeyword) <> 0 Then

                    ' Do nothing
                    ' Note: Enum declarations do not have data type specified because they are always
                    ' of data type long.
                    ' We check for Enum keyword because it may be preceded by the Public or Private
                    ' keywords which can be used to declare a variable.  If we did not check for
                    ' the Enum keyword now, we would have to verify that it did not exist later
                    ' when looking for variable declarations.

                ' Test if logical line is variable declaration
                ElseIf ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Public", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Private", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Dim", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Const", tttcKeyword) <> 0 Then

                    ' Split search variables string into search variables array containing one search variable per element
                    astrSearchVariables() = Split(strSearchLogicalLine, ",", -1, vbBinaryCompare)
                    ' Cycle through all search variables in array
                    For lngSearchVariableIndex = 0 To UBound(astrSearchVariables)
                        ' Test if variable data type is missing
                        If ToolGetTokenPosition(astrSearchVariables(lngSearchVariableIndex), 1, True, "As", tttcKeyword) = 0 Then
                            fDeclarationWithoutDefinedDataType = True
                            Exit For
                        End If
                    Next

                End If

                If lngCompoundStatementPhysicalLineCount = 0 Then

                    If fDeclarationWithoutDefinedDataType Then
                        Debug.Print objVBComponent.Name & vbCrLf & strLogicalLine & vbCrLf
                        strLog = strLog & objVBComponent.Name & vbCrLf & strLogicalLine & vbCrLf & vbCrLf
                    End If

                    ' Set current line to last physical line of logical line so next
                    ' loop iteration will check next physical line after logical line
                    lngLine = lngLine + lngLogicalLinePhysicalLineCount - 1

                Else

                    If fDeclarationWithoutDefinedDataType Then
                        Debug.Print objVBComponent.Name & vbCrLf & StrCompoundStatementPhysicalLines & vbCrLf
                        strLog = strLog & objVBComponent.Name & vbCrLf & StrCompoundStatementPhysicalLines & vbCrLf & vbCrLf
                    End If

                    ' Set current line to last physical line of compound statement so next
                    ' loop iteration will check next physical line after compound statement
                    lngLine = lngLine + lngCompoundStatementPhysicalLineCount - 1

                End If

            Next

            ' Cycle through all non declaration lines of code
            For lngLine = .CountOfDeclarationLines + 1 To .CountOfLines

                ' Reset declaration without defined data type flag
                fDeclarationWithoutDefinedDataType = False

                ' Get logical line and count of physical lines in logical line
                strLogicalLine = ToolGetLogicalLine(objVBComponent.CodeModule, lngLine, lngLogicalLinePhysicalLineCount)

                ' Get search logical line for use when searching for token positions
                strSearchLogicalLine = ToolGetTokenSearchLogicalLine(strLogicalLine)

                ' Test if logical line is sub, function or property declaration
                If (ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Sub", tttcKeyword) <> 0 And _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Exit Sub", tttcKeyword) = 0) Or _
                    (ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Function", tttcKeyword) <> 0 And _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Exit Function", tttcKeyword) = 0) Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Property Get", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Property Let", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Property Set", tttcKeyword) <> 0 Then

                    ' Get arguments start and end positions
                    lngArgumentsStartPosition = ToolGetTokenPosition(strSearchLogicalLine, 1, True, "(", tttcSeparator)
                    lngArgumentsStartPosition = lngArgumentsStartPosition + 1
                    lngArgumentsEndPosition = ToolGetTokenPosition(strSearchLogicalLine, -1, False, ")", tttcSeparator)
                    lngArgumentsEndPosition = lngArgumentsEndPosition - 1
                    ' Verify one or more arguments exists or procedure returns and array
                    If lngArgumentsEndPosition > lngArgumentsStartPosition Then
                        ' Test if procedure returns and array
                        If Mid$(strSearchLogicalLine, lngArgumentsEndPosition, 1) = "(" Then
                            ' Procedure returns array therefore need to find next ')'
                            ' searching in reverse to find arguments end position
                            lngArgumentsEndPosition = ToolGetTokenPosition(strSearchLogicalLine, lngArgumentsEndPosition, False, ")", tttcSeparator)
                            lngArgumentsEndPosition = lngArgumentsEndPosition - 1
                        End If
                    End If
                    ' Verify one or more arguments exists
                    If lngArgumentsEndPosition > lngArgumentsStartPosition Then
                        ' Get search arguments string
                        strSearchArguments = Mid$(strSearchLogicalLine, lngArgumentsStartPosition, lngArgumentsEndPosition - lngArgumentsStartPosition + 1)
                        ' Split search arguments string into search arguments array containing one search argument per element
                        astrSearchArguments() = Split(strSearchArguments, ",", -1, vbBinaryCompare)
                        ' Cycle through all search arguments in array
                        For lngSearchArgumentIndex = 0 To UBound(astrSearchArguments)
                            ' Test if argument data type is missing
                            If ToolGetTokenPosition(astrSearchArguments(lngSearchArgumentIndex), 1, True, "As", tttcKeyword) = 0 Then
                                fDeclarationWithoutDefinedDataType = True
                                Exit For
                            End If
                        Next
                        ' Verify we have not already determined that data type is missing for at least one argument
                        If Not fDeclarationWithoutDefinedDataType Then
                            ' Test if declaration is for procedure with return data type
                            If ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Function", tttcKeyword) <> 0 Or _
                                ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Get", tttcKeyword) <> 0 Then
                                ' Get return data type start position
                                lngReturnDataTypeStartPosition = lngArgumentsEndPosition + 2
                                ' Test if return data type is missing
                                If ToolGetTokenPosition(strSearchLogicalLine, lngReturnDataTypeStartPosition, True, "As", tttcKeyword) = 0 Then
                                    fDeclarationWithoutDefinedDataType = True
                                End If
                            End If
                        End If
                    End If

                ' Test if logical line is variable declaration
                ElseIf ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Dim", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Static", tttcKeyword) <> 0 Or _
                    ToolGetTokenPosition(strSearchLogicalLine, 1, True, "Const", tttcKeyword) <> 0 Then

                    ' Split search variables string into search variables array containing one search variable per element
                    astrSearchVariables() = Split(strSearchLogicalLine, ",", -1, vbBinaryCompare)
                    ' Cycle through all search variables in array
                    For lngSearchVariableIndex = 0 To UBound(astrSearchVariables)
                        ' Test if variable data type is missing
                        If ToolGetTokenPosition(astrSearchVariables(lngSearchVariableIndex), 1, True, "As", tttcKeyword) = 0 Then
                            fDeclarationWithoutDefinedDataType = True
                            Exit For
                        End If
                    Next

                End If

                If fDeclarationWithoutDefinedDataType Then
                    Debug.Print objVBComponent.Name & vbCrLf & strLogicalLine & vbCrLf
                    strLog = strLog & objVBComponent.Name & vbCrLf & strLogicalLine & vbCrLf & vbCrLf
                End If

                ' Set current line to last physical line of logical line so next
                ' loop iteration will check next physical line after logical line
                lngLine = lngLine + lngLogicalLinePhysicalLineCount - 1

            Next

        End With
    Next
    
    ToolSaveStringToFile Environ$("TEMP") & "\Desktop\Log.txt", strLog

End Sub

Private Function ToolGetCharacterCode( _
    ByRef strCharacters As String, _
    ByRef lngCharacterPosition As Long) As Integer

On Error Resume Next

    Dim Property As Long

    ' Purpose:
    ' Return integer for character code of first byte of a string character.
    ' Return 0 if position is before start or after end of input string.
    ' This make sense because 0 is the character code for a Null character
    ' which is used to mark the end of a string.  Null characters are not
    ' allowed within a VBA string therefore by returning a character code
    ' of 0 we are saying no character exists at that position.

    ' Assumptions:
    ' All character codes will be <= 255 when working within the VBA IDE.

    ToolGetCharacterCode = AscB(Mid$(strCharacters, lngCharacterPosition, 1))

End Function

Private Function ToolGetLogicalLine( _
    ByRef objCodeModule As Object, _
    ByRef lngPhysicalStartLine As Long, _
    Optional ByRef lngPhysicalLineCount As Long) As String

    ' Purpose:
    ' Procedure used to get all physical lines of code that make up a single logical line.
    ' A logical line ends with a line terminator that is not preceded by a line continuation.

    ' Restrictions:
    ' Multiple logical lines (i.e. statements) can be added to a single physical line using the ':' character.
    ' This procedure cannot handle multiple logical lines (i.e. statements) on the same physical line.

    ' Arguments:
    ' objCodeModule: In
    ' lngPhysicalStartLine: In
    ' lngPhysicalLineCount: Out

    lngPhysicalLineCount = 1
    Do
        ' Increment line count if line continuation character exists
        If Right$(objCodeModule.Lines(lngPhysicalStartLine + lngPhysicalLineCount - 1, 1), 2) = " _" Then
            lngPhysicalLineCount = lngPhysicalLineCount + 1
        Else
            Exit Do
        End If
    Loop

    ToolGetLogicalLine = objCodeModule.Lines(lngPhysicalStartLine, lngPhysicalLineCount)

End Function

Private Sub ToolDeleteLogicalLine( _
    ByRef objCodeModule As Object, _
    ByRef lngPhysicalStartLine As Long, _
    Optional ByRef lngPhysicalLineCount As Long)

    ' Purpose:
    ' Procedure used to delete all physical lines of code that make up a single logical line.
    ' A logical line ends with a line terminator that is not preceded by a line continuation.

    ' Restrictions:
    ' Multiple logical lines (i.e. statements) can be added to a single physical line using the ':' character.
    ' This procedure cannot handle multiple logical lines (i.e. statements) on the same physical line.

    ' Arguments:
    ' objCodeModule: In
    ' lngPhysicalStartLine: In
    ' lngPhysicalLineCount: Out

    lngPhysicalLineCount = 1
    Do
        ' Increment line count if line continuation character exists
        If Right$(objCodeModule.Lines(lngPhysicalStartLine + lngPhysicalLineCount - 1, 1), 2) = " _" Then
            lngPhysicalLineCount = lngPhysicalLineCount + 1
        Else
            Exit Do
        End If
    Loop

    objCodeModule.DeleteLines lngPhysicalStartLine, lngPhysicalLineCount

End Sub

Private Function ToolCharacterIsWhiteSpace( _
    ByRef strLogicalLine As String, _
    ByRef lngCharacterPosition As Long) As Boolean

    ' Purpose:
    ' Return true if character at position is white-space.

    ' Assumptions:
    ' All character codes will be <= 255 when working within the VBA IDE.

    ' Initialize static variables if not already intialized
    Static sfInitialized As Boolean
    If Not sfInitialized Then

        ' Build array that returns true when index is character code
        ' of white-space-character, or line-terminator character
        Static sabyteWhiteSpaceCharacterCodes(255) As Boolean

        ' Null
        ' Note*: Although Null characters should not exist within a string
        ' if they do exist it makes sense to treat them as white-space.
        ' Also, the ToolGetCharacterCode procedure will return Null if a
        ' position is before start of after end of string.
        sabyteWhiteSpaceCharacterCodes(0) = True  ' Null

        ' White-space-characters
        sabyteWhiteSpaceCharacterCodes(9) = True  ' tab
        sabyteWhiteSpaceCharacterCodes(25) = True ' end of medium
        sabyteWhiteSpaceCharacterCodes(32) = True ' space

        ' Line-terminators
        sabyteWhiteSpaceCharacterCodes(10) = True ' line feed
        sabyteWhiteSpaceCharacterCodes(13) = True ' carriage return

        ' Flag that static variables have been initialized
        sfInitialized = True

    End If

    ' Return true if character at position is white-space
    ' 1. If character is white-space-character, it is white-space.
    ' 2. If character is part of line-continuation, it is white-space.
    '    If character is line-terminator character, it must be part
    '    of a line-continuation.  If character is underscore character,
    '    it may part of a line-continuation.  An underscore is part of
    '    a line-continuation if it is left and right delimited with
    '    white-space-characters and/or line-terminator characters.
    Dim intCharacterCode As Integer
    intCharacterCode = ToolGetCharacterCode(strLogicalLine, lngCharacterPosition)
    If sabyteWhiteSpaceCharacterCodes(intCharacterCode) Then
        ToolCharacterIsWhiteSpace = True
    Else
        If intCharacterCode = 95 Then ' 95 is character code for underscore
            Dim intLeftCharacterCode As Integer
            intLeftCharacterCode = ToolGetCharacterCode(strLogicalLine, lngCharacterPosition - 1)
            If sabyteWhiteSpaceCharacterCodes(intLeftCharacterCode) Then
                Dim intRightCharacterCode As Integer
                intRightCharacterCode = ToolGetCharacterCode(strLogicalLine, lngCharacterPosition + 1)
                If sabyteWhiteSpaceCharacterCodes(intRightCharacterCode) Then
                    ToolCharacterIsWhiteSpace = True
                End If
            End If
        End If
    End If

End Function

Private Function ToolCharacterIsSpecialToken( _
    ByRef strLogicalLine As String, _
    ByRef lngCharacterPosition As Long) As Boolean

    ' Purpose:
    ' Return true if character at position is special-token.

    ' Assumptions:
    ' All character codes will be <= 255 when working within the VBA IDE.

    ' Initialize static variables if not already intialized
    Static sfInitialized As Boolean
    If Not sfInitialized Then

        ' Build array that returns true when index is character code
        ' of special-token character
        Static sabyteSpecialTokenCharacterCodes(255) As Boolean

        ' Separator special-tokens
        sabyteSpecialTokenCharacterCodes(40) = True ' (
        sabyteSpecialTokenCharacterCodes(41) = True ' )
        sabyteSpecialTokenCharacterCodes(44) = True ' ,
        sabyteSpecialTokenCharacterCodes(46) = True ' .
        sabyteSpecialTokenCharacterCodes(33) = True ' !
        sabyteSpecialTokenCharacterCodes(35) = True ' #
        sabyteSpecialTokenCharacterCodes(58) = True ' :
        sabyteSpecialTokenCharacterCodes(63) = True ' ?
        ' Arithmetic operator special-tokens
        sabyteSpecialTokenCharacterCodes(94) = True ' ^
        sabyteSpecialTokenCharacterCodes(42) = True ' *
        sabyteSpecialTokenCharacterCodes(47) = True ' /
        sabyteSpecialTokenCharacterCodes(92) = True ' \
        sabyteSpecialTokenCharacterCodes(43) = True ' +
        sabyteSpecialTokenCharacterCodes(45) = True ' -
        ' Comparison operator special-tokens
        sabyteSpecialTokenCharacterCodes(60) = True ' <
        sabyteSpecialTokenCharacterCodes(61) = True ' =
        sabyteSpecialTokenCharacterCodes(62) = True ' >
        ' Concatenation operator special-tokens
        sabyteSpecialTokenCharacterCodes(38) = True ' &
        'sabyteSpecialTokenCharacterCodes(43) = True ' Already flagged as arithmetic operator
        ' End of statement special-token
        sabyteSpecialTokenCharacterCodes(59) = True ' ;
        ' Start of comment special-token
        sabyteSpecialTokenCharacterCodes(39) = True ' ' (i.e. single-quote)

        ' Flag that static variables have been initialized
        sfInitialized = True

    End If

    ' Return true if character at position is special-token
    Dim intCharacterCode As Integer
    intCharacterCode = ToolGetCharacterCode(strLogicalLine, lngCharacterPosition)
    If sabyteSpecialTokenCharacterCodes(intCharacterCode) Then
        ToolCharacterIsSpecialToken = True
    End If

End Function

Private Function ToolTokenIsDelimited( _
    ByRef strLogicalLine As String, _
    ByRef lngTokenPosition As Long, _
    ByRef lngTokenLength As Long) As Boolean

    ' Purpose:
    ' Return true if token is left and right delimited by any character that
    ' is not a letter, number, or underscore and has a character code <=127.

    ' Valid delimiter characters were found by attempting to right delimit the
    ' Rem keyword with all characters that have a character code from 0 to 255.
    ' The Rem keyword was chosen because it starts a comment therefore any
    ' character after Rem keyword should be valid as long as it delimits the
    ' Rem keyword.  Test results indicated:
    ' 1. No characters that have a character code > 127 can be used as a delimiter.
    ' 2. All characters that have a character code =< 127 can be used as a delimiter
    '    except numbers, letters and the underscore character.
    ' This is consistent with the following identifier and keyword naming rules:
    ' 1. Must begin with a letter.
    ' 2. Must contain only letters, numbers, and underscores (no other symbols).
    ' 3. Cannot be more than 255 characters long.
    ' The fact that a character can be used to delimit an identifer or keyword
    ' token does not mean that character is valid.  It only means that the
    ' character can be used to reliable indicate the start and end of an
    ' identifier or keyword token.

    ' Assumptions:
    ' All character codes will be <= 255 when working within the VBA IDE.

    ' Initialize static variables if not already intialized
    Static sfInitialized As Boolean
    If Not sfInitialized Then

        ' Build array that returns true when index is character code
        ' of character that is a valid token delimiter character.
        Static sabyteDelimiterCharacterCodes(255) As Boolean
        Dim lngDelimiterCharacterCodeIndex As Long
        For lngDelimiterCharacterCodeIndex = 0 To 47
            sabyteDelimiterCharacterCodes(lngDelimiterCharacterCodeIndex) = True
        Next
        ' Skip character codes 48 through 57 for characters 0 through 9
        For lngDelimiterCharacterCodeIndex = 58 To 63
            sabyteDelimiterCharacterCodes(lngDelimiterCharacterCodeIndex) = True
        Next
        ' Skip character codes 64 through 90 for characters A through Z
        For lngDelimiterCharacterCodeIndex = 91 To 94
            sabyteDelimiterCharacterCodes(lngDelimiterCharacterCodeIndex) = True
        Next
        ' Skip character code 95 for character '_'
        For lngDelimiterCharacterCodeIndex = 96 To 97
            sabyteDelimiterCharacterCodes(lngDelimiterCharacterCodeIndex) = True
        Next
        ' Skip character codes 97 through 122 for characters a through z
        For lngDelimiterCharacterCodeIndex = 122 To 127
            sabyteDelimiterCharacterCodes(lngDelimiterCharacterCodeIndex) = True
        Next

        ' Flag that static variables have been initialized
        sfInitialized = True

    End If

    ' Return true if token at position is left and right delimited by
    ' white-space, special-token or the start/end of the logical line
    ' 1. If character is white-space-character, it is white-space
    '    therefore delimits token.
    ' 2. If character is part of line-continuation, it is white-space
    '    therefore delimits token.  If character is line-terminator
    '    character, it must be part of a line-continuation.  If
    '    character is an underscore character, it must not be part
    '    of a line-continuation because a line-continuation underscore
    '    will be left and right delimited with white-space-characters
    '    and/or line-terminator characters.
    ' 3. If character is a special-token, it delimits token.
    ' 4. If character is Null, position is before start or after end
    '    of logical line therefore delimits token.
    Dim intLeftCharacterCode As Integer
    intLeftCharacterCode = ToolGetCharacterCode(strLogicalLine, lngTokenPosition - 1)
    If sabyteDelimiterCharacterCodes(intLeftCharacterCode) Then
        Dim intRightCharacterCode As Integer
        intRightCharacterCode = ToolGetCharacterCode(strLogicalLine, lngTokenPosition + lngTokenLength)
        If sabyteDelimiterCharacterCodes(intRightCharacterCode) Then
            ToolTokenIsDelimited = True
        End If
    End If

End Function

Private Function ToolTokenIsDelimitedIfNeeded( _
    ByRef strLogicalLine As String, _
    ByRef lngTokenPosition As Long, _
    ByRef lngTokenLength As Long, _
    ByRef tttTokenType As tttcToolTokenType) As Boolean

    ' Purpose:
    ' Return true if token type does need to be delmited and token is
    ' left and right delimited by any character that is not a letter,
    ' number, or underscore and has a character code <=127.
    ' Return true if token type does not need to be delimited.

    Select Case tttTokenType
        Case tttcIdentifier, tttcKeyword
            ' Identifiers and keywords must be delimited
            ToolTokenIsDelimitedIfNeeded = ToolTokenIsDelimited(strLogicalLine, lngTokenPosition, lngTokenLength)
        Case tttcNumberLiteral, tttcDateLiteral, tttcStringLiteral, tttcSeparator
            ' Number literals, date Literals, string literals, and separators do not need to be delimited
            ToolTokenIsDelimitedIfNeeded = True
        Case tttcOperator
            If ToolCharacterIsSpecialToken(strLogicalLine, lngTokenPosition) Then
                ' Operators that are special-tokens do not need to be delimited
                ToolTokenIsDelimitedIfNeeded = True
            Else
                ' Operators that are not special-tokens such as 'Mod', 'Like', and 'Not'
                ' are special types of identifiers therefore must be delimited.
                ToolTokenIsDelimitedIfNeeded = ToolTokenIsDelimited(strLogicalLine, lngTokenPosition, lngTokenLength)
            End If
    End Select

End Function

Private Function ToolGetFirstTokenPosition( _
    ByRef strLogicalLine As String, _
    ByRef lngSearchStart As Long, _
    ByRef fSearchForward As Boolean) As Long

    ' Purpose:
    ' Return position of first token found (i.e. first character that is not white-space).

    Dim lngCharacterPosition As Long
    If fSearchForward Then
        For lngCharacterPosition = lngSearchStart To Len(strLogicalLine)
            If Not ToolCharacterIsWhiteSpace(strLogicalLine, lngCharacterPosition) Then
                ToolGetFirstTokenPosition = lngCharacterPosition
                Exit For
            End If
        Next
    Else
        ' Note: -1 is used to indicate that search should start at the end of the string
        If lngSearchStart = -1 Then _
            lngSearchStart = Len(strLogicalLine)
        For lngCharacterPosition = lngSearchStart To 1 Step -1
            If Not ToolCharacterIsWhiteSpace(strLogicalLine, lngCharacterPosition) Then
                ToolGetFirstTokenPosition = lngCharacterPosition
                Exit For
            End If
        Next
    End If

End Function

Private Function ToolParseLogicalLine( _
    ByRef strLogicalLine As String, _
    Optional ByRef fOverwriteCommentWithWhiteSpace As Boolean, _
    Optional ByRef fOverwriteStringLiteralsWithWhiteSpace As Boolean, _
    Optional ByRef fOverwriteDateLiteralsWithWhiteSpace As Boolean, _
    Optional ByRef fOverwriteNumberLiteralsWithWhiteSpace As Boolean) As String

    ' Purpose:
    ' Return parsed logical line.  Comment and literals can be overwritten
    ' with white space making it eash to search for tokens without having
    ' to handle the case were a token being searched for is found inside
    ' a comment or literal.

    ' Restrictions:
    ' Rem keyword can be used to create a logical line that only contains a comment.
    ' Only a single-quote character can be used to add a comment to the end of a
    ' logical line following a statement.  This procedure does not blank comments
    ' defined using Rem keyword.  Procedure could easily be updated to do so but
    ' Rem keyword is never used in practice.

    ' Assumptions:
    ' All character codes will be <= 255 when working within the VBA IDE.
    ' All literals will have a valid syntax.  Literals will invalid syntax
    ' will not be parsed correctly.

    ' Initialize static variables if not already intialized
    Static sfInitialized As Boolean
    If Not sfInitialized Then

        ' Build array that returns true when index is character code
        ' of valid hexadecimal symbol character
        Static sabyteHexadecimalSymbolCharacterCodes(255) As Boolean
        Dim lngHexadecimalLiteralCharacterCodeIndex As Long
        For lngHexadecimalLiteralCharacterCodeIndex = 48 To 57  '  0 through 9
            sabyteHexadecimalSymbolCharacterCodes(lngHexadecimalLiteralCharacterCodeIndex) = True
        Next
        For lngHexadecimalLiteralCharacterCodeIndex = 65 To 70  '  A through F
            sabyteHexadecimalSymbolCharacterCodes(lngHexadecimalLiteralCharacterCodeIndex) = True
        Next
        For lngHexadecimalLiteralCharacterCodeIndex = 97 To 102 '  a through f
            sabyteHexadecimalSymbolCharacterCodes(lngHexadecimalLiteralCharacterCodeIndex) = True
        Next

        ' Build array that returns true when index is character code
        ' of valid octal symbol character
        Static sabyteOctalSymbolCharacterCodes(255) As Boolean
        Dim lngOctalLiteralCharacterCodeIndex As Long
        For lngOctalLiteralCharacterCodeIndex = 48 To 55 ' 0 through 7
            sabyteOctalSymbolCharacterCodes(lngOctalLiteralCharacterCodeIndex) = True
        Next

        ' Build array that returns true when index is character code
        ' of valid decimal symbol character
        Static sabyteDecimalSymbolCharacterCodes(255) As Boolean
        Dim lngDecimalLiteralCharacterCodeIndex As Long
        For lngDecimalLiteralCharacterCodeIndex = 48 To 57 ' 0 through 9
            sabyteDecimalSymbolCharacterCodes(lngDecimalLiteralCharacterCodeIndex) = True
        Next

        ' Build array that returns true when index is character code
        ' of letter or underscore character
        Static sabyteLetterOrUnderscoreCharacterCodes(255) As Boolean
        Dim lngLetterOrUnderscoreCharacterCodeIndex As Long
        For lngLetterOrUnderscoreCharacterCodeIndex = 65 To 90 ' A through Z
            sabyteLetterOrUnderscoreCharacterCodes(lngLetterOrUnderscoreCharacterCodeIndex) = True
        Next
        For lngLetterOrUnderscoreCharacterCodeIndex = 97 To 122 ' a through z
            sabyteLetterOrUnderscoreCharacterCodes(lngLetterOrUnderscoreCharacterCodeIndex) = True
        Next
        sabyteLetterOrUnderscoreCharacterCodes(95) = True ' _

        ' Build array that returns true when index is character code
        ' of valid type declaration character
        Static sabyteTypedDeclarationCharacterCodes(255) As Boolean
        sabyteTypedDeclarationCharacterCodes(33) = True ' !
        sabyteTypedDeclarationCharacterCodes(35) = True ' #
        sabyteTypedDeclarationCharacterCodes(37) = True ' %
        sabyteTypedDeclarationCharacterCodes(38) = True ' &
        sabyteTypedDeclarationCharacterCodes(33) = True ' &
        sabyteTypedDeclarationCharacterCodes(64) = True ' @

        ' Flag that static variables have been initialized
        sfInitialized = True

    End If

    ' Initialize return logical line with copy of logical line passed to procedure
    Dim strReturnLogicalLine As String
    strReturnLogicalLine = strLogicalLine

    ' Get length of logical line
    Dim lngLogicalLineLength As Long
    lngLogicalLineLength = Len(strLogicalLine)

    ' Cycle through all characters in logical line blanking literals and comments
    ' as we go.
    Dim lngCharacterPosition As Long
    For lngCharacterPosition = 1 To lngLogicalLineLength

        Dim intCharacterCode As Integer
        intCharacterCode = ToolGetCharacterCode(strLogicalLine, lngCharacterPosition)
        Select Case intCharacterCode

            Case 39 ' single-quote

                ' Comments start with a ' character and last until the end of the logical line
                ' The ' character may only be used as follows:
                ' 1. Within a comment.
                ' 2. Within a string literal.
                ' 3. Within a date literal.
                ' 4. To mark the start of a comment.

                ' Overwrite comment with white space if enabled
                If fOverwriteCommentWithWhiteSpace Then

                    ' Get length of comment
                    Dim lngCommentLength As Long
                    lngCommentLength = lngLogicalLineLength - lngCharacterPosition + 1

                    ' Overwrite comment with white space
                    Mid$(strReturnLogicalLine, lngCharacterPosition, lngCommentLength) = Space(lngCommentLength)

                End If

                ' Comments last until the end of the logical line therefore exit for loop
                Exit For

            Case 34 ' "

                ' String literals must be enclosed in " characters.
                ' The " character may only be used as follows:
                ' 1. Within a comment.
                ' 2. Within a string literal.
                ' 3. To enclose a string literal.

                ' Find end of string literal
                Dim lngStringLiteralCharacterPosition As Long
                For lngStringLiteralCharacterPosition = lngCharacterPosition + 1 To lngLogicalLineLength
                    ' Test if character is " indicating end of string literal
                    If ToolGetCharacterCode(strLogicalLine, lngStringLiteralCharacterPosition) = 34 Then
                        ' Verify there are not two " characters in a row indicating
                        ' an escaped double quote within a string literal
                        If Not ToolGetCharacterCode(strLogicalLine, lngCharacterPosition + 1) = 34 Then
                            ' We have found the end of the string literal, exit for loop
                            Exit For
                        End If
                    End If
                Next

                ' Get length of string literal
                Dim lngStringLiteralLength As Long
                lngStringLiteralLength = lngStringLiteralCharacterPosition - lngCharacterPosition + 1

                ' Overwrite string literal with white space if enabled
                If fOverwriteStringLiteralsWithWhiteSpace Then
                    Mid$(strReturnLogicalLine, lngCharacterPosition, lngStringLiteralLength) = Space(lngStringLiteralLength)
                End If

                ' Set character position to last character of string literal so that
                ' next loop iteration will check first character after string literal
                lngCharacterPosition = lngCharacterPosition + lngStringLiteralLength - 1

            Case 35 ' #

                ' Date literals must be enclosed in # characters.
                ' The # character may only be used as follows:
                ' 1. Within a comment.
                ' 2. Within a string literal.
                ' 3. To enclose a date literal.
                ' 4. As type declaration character at the end of a number literal
                '    to indicate number data type of double.

                ' Find end of date literal
                Dim lngDateLiteralCharacterPosition As Long
                For lngDateLiteralCharacterPosition = lngCharacterPosition + 1 To lngLogicalLineLength
                    ' Test if character is # indicating end of date literal
                    If ToolGetCharacterCode(strLogicalLine, lngDateLiteralCharacterPosition) = 35 Then
                        ' We have found the end of the date literal, exit for loop
                        Exit For
                    End If
                Next

                ' Get length of date literal
                Dim lngDateLiteralLength As Long
                lngDateLiteralLength = lngDateLiteralCharacterPosition - lngCharacterPosition + 1

                ' Overwrite date literal with white space if enabled
                If fOverwriteDateLiteralsWithWhiteSpace Then
                    Mid$(strReturnLogicalLine, lngCharacterPosition, lngDateLiteralLength) = Space(lngDateLiteralLength)
                End If

                ' Set character position to last character of date literal so that
                ' next loop iteration will check first character after date literal
                lngCharacterPosition = lngCharacterPosition + lngDateLiteralLength - 1

            Case Else

                ' Hexadecimal literals start with &H or &h and end
                ' with 0 through A or a type declaration character.
                ' Octal literals start with &O or &o and end
                ' with 0 through 7 or a type declaration character.
                ' Decimal literals start with 0 through 9 or '.' and end
                ' with 0 through 9, '.' or type declaration character.

                ' Test for '&' character because it may indicate the start of a hexadecimal or octal literal
                If intCharacterCode = 38 Then

                    Select Case ToolGetCharacterCode(strLogicalLine, lngCharacterPosition + 1)

                        ' Test if first character after '&' character is a 'H' or 'h'
                        ' to determine if we found the start of a hexadecimal literal
                        Case 72, 104 ' H, h

                            ' Find position 1 character after end of hexadecimal literal
                            Dim lngHexadecimalLiteralCharacterPosition As Long
                            For lngHexadecimalLiteralCharacterPosition = lngCharacterPosition + 2 To lngLogicalLineLength + 1
                                Dim intHexadecimalLiteralCharacterCode As Integer
                                intHexadecimalLiteralCharacterCode = ToolGetCharacterCode(strLogicalLine, lngHexadecimalLiteralCharacterPosition)
                                If Not sabyteHexadecimalSymbolCharacterCodes(intHexadecimalLiteralCharacterCode) Then
                                    If sabyteTypedDeclarationCharacterCodes(intHexadecimalLiteralCharacterCode) Then
                                        lngHexadecimalLiteralCharacterPosition = lngHexadecimalLiteralCharacterPosition + 1
                                    End If
                                    Exit For
                                End If
                            Next

                            ' Get length of hexadecimal literal
                            Dim lngHexadecimalLiteralLength As Long
                            lngHexadecimalLiteralLength = lngHexadecimalLiteralCharacterPosition - lngCharacterPosition

                            ' Overwrite hexadecimal literal with white space if enabled
                            If fOverwriteNumberLiteralsWithWhiteSpace Then
                                Mid$(strReturnLogicalLine, lngCharacterPosition, lngHexadecimalLiteralLength) = Space(lngHexadecimalLiteralLength)
                            End If

                            ' Set character position to last character of hexadecimal literal so that
                            ' next loop iteration will check first character after hexadecimal literal
                            lngCharacterPosition = lngCharacterPosition + lngHexadecimalLiteralLength - 1

                        ' Test if first character after '&' character is a 'O' or 'o'
                        ' to determine if we found the start of an octal literal
                        Case 79, 111 ' O, o

                            ' Find position 1 character after end of octal literal
                            Dim lngOctalLiteralCharacterPosition As Long
                            For lngOctalLiteralCharacterPosition = lngCharacterPosition + 2 To lngLogicalLineLength + 1
                                Dim intOctalLiteralCharacterCode As Integer
                                intOctalLiteralCharacterCode = ToolGetCharacterCode(strLogicalLine, lngOctalLiteralCharacterPosition)
                                If Not sabyteOctalSymbolCharacterCodes(intOctalLiteralCharacterCode) Then
                                    If sabyteTypedDeclarationCharacterCodes(intOctalLiteralCharacterCode) Then
                                        lngOctalLiteralCharacterPosition = lngOctalLiteralCharacterPosition + 1
                                    End If
                                    Exit For
                                End If
                            Next

                            ' Get length of octal literal
                            Dim lngOctalLiteralLength As Long
                            lngOctalLiteralLength = lngOctalLiteralCharacterPosition - lngCharacterPosition

                            ' Overwrite octal literal with white space if enabled
                            If fOverwriteNumberLiteralsWithWhiteSpace Then
                                Mid$(strReturnLogicalLine, lngCharacterPosition, lngOctalLiteralLength) = Space(lngOctalLiteralLength)
                            End If

                            ' Set character position to last character of octal literal so that
                            ' next loop iteration will check first character after octal literal
                            lngCharacterPosition = lngCharacterPosition + lngOctalLiteralLength - 1

                    End Select

                ' Test for decimal symbol character because it indicates the start of a sequence of
                ' one or more decimal symbols within an identifier or the start of a decimal literal
                ElseIf sabyteDecimalSymbolCharacterCodes(intCharacterCode) Then

                    ' Test if first character before decimal symbol is a letter or an underscore
                    ' to determine if we have found the start of a decimal symbol sequence within
                    ' an identifier or the start of a decimal literal.
                    ' Note*: Identifier naming rules ensure that their names do not start with decimal symbol.
                    If sabyteLetterOrUnderscoreCharacterCodes(ToolGetCharacterCode(strLogicalLine, lngCharacterPosition - 1)) Then

                        ' Find position 1 character after end of decimal symbol sequence within an identifier
                        Dim lngDecimalSymbolSequenceCharacterPosition As Long
                        For lngDecimalSymbolSequenceCharacterPosition = lngCharacterPosition + 1 To lngLogicalLineLength + 1
                            Dim intDecimalSymbolSequenceCharacterCode As Integer
                            intDecimalSymbolSequenceCharacterCode = ToolGetCharacterCode(strLogicalLine, lngDecimalSymbolSequenceCharacterPosition)
                            If Not sabyteDecimalSymbolCharacterCodes(intDecimalSymbolSequenceCharacterCode) Then
                                If sabyteTypedDeclarationCharacterCodes(intDecimalSymbolSequenceCharacterCode) Then
                                    lngDecimalSymbolSequenceCharacterPosition = lngDecimalSymbolSequenceCharacterPosition + 1
                                End If
                                Exit For
                            End If
                        Next

                        ' Get length of decimal symbol sequence within an identifier
                        Dim lngDecimalSymbolSequenceLength As Long
                        lngDecimalSymbolSequenceLength = lngDecimalSymbolSequenceCharacterPosition - lngCharacterPosition

                        ' Set character position to last character of decimal symbol sequence within an identifier so that
                        ' next loop iteration will check first character after decimal symbol sequence within an identifier
                        lngCharacterPosition = lngCharacterPosition + lngDecimalSymbolSequenceLength - 1

                    Else

                        ' Find position 1 character after end of decimal literal
                        Dim lngDecimalLiteralCharacterPosition As Long
                        For lngDecimalLiteralCharacterPosition = lngCharacterPosition + 1 To lngLogicalLineLength + 1
                            Dim intDecimalLiteralCharacterCode As Integer
                            intDecimalLiteralCharacterCode = ToolGetCharacterCode(strLogicalLine, lngDecimalLiteralCharacterPosition)
                            If Not sabyteDecimalSymbolCharacterCodes(intDecimalLiteralCharacterCode) Then
                                ' Note: A decimal point character may occur once within a floating point
                                ' decimal number.  We don't need to verify it only occurs once because
                                ' this procedure assumes all literals will have a valid syntax.  If we
                                ' included the decimal point character in list of decimal symbols we
                                ' could incorrectly detect a decimal literal because a decimal point
                                ' can also be a delimiter (e.g. 'frm1.txtControl1.Value').
                                If Not intDecimalLiteralCharacterCode = 46 Then ' Not decimal point character
                                    If sabyteTypedDeclarationCharacterCodes(intDecimalLiteralCharacterCode) Then
                                        lngDecimalLiteralCharacterPosition = lngDecimalLiteralCharacterPosition + 1
                                    End If
                                    Exit For
                                End If
                            End If
                        Next

                        ' Get length of decimal literal
                        Dim lngDecimalLiteralLength As Long
                        lngDecimalLiteralLength = lngDecimalLiteralCharacterPosition - lngCharacterPosition

                        ' Overwrite decimal literal with white space if enabled
                        If fOverwriteNumberLiteralsWithWhiteSpace Then
                            Mid$(strReturnLogicalLine, lngCharacterPosition, lngDecimalLiteralLength) = Space(lngDecimalLiteralLength)
                        End If

                        ' Set character position to last character of decimal literal so that
                        ' next loop iteration will check first character after decimal literal
                        lngCharacterPosition = lngCharacterPosition + lngDecimalLiteralLength - 1

                    End If

                End If

        End Select

    Next

    ' Return parsed logical line
    ToolParseLogicalLine = strReturnLogicalLine

End Function

Private Function ToolGetTokenSearchLogicalLine( _
    ByRef strLogicalLine As String) As String

    ' Purpose:
    ' Return string containing logical line with comment and literals overwritten with
    ' white space to prevent finding a token within a comment or literal when searching.

    ' Restrictions:
    ' This procedure cannot be used if it is desired to search within a comment or for
    ' a literal token because this procedures overwrites them with white space.

    ToolGetTokenSearchLogicalLine = ToolParseLogicalLine(strLogicalLine, True, True, True, True)

End Function

Private Function ToolGetTokenPosition( _
    ByRef strSearchLogicalLine As String, _
    ByRef lngSearchStart As Long, _
    ByRef fSearchForward As Boolean, _
    ByRef strToken As String, _
    ByRef tttTokenType As tttcToolTokenType) As Long

    ' Purpose:
    ' Return position of token.

    ' Assumptions:
    ' The strSearchLogicalLine argument passed to this procedure has already had
    ' comments and literals removed that could prevent correctly finding a token.
    ' Note*: This procedure expects the strSearchLogicalLine argument to already
    ' have comments and literals removed for performance reasons.
    ' The ToolGetTokenSearchLogicalLine procedure calls the ToolParseLogicalLine
    ' procedure to remove comments and literals from a logical line.
    ' The ToolParseLogicalLine procedures checks all characters of a logical line
    ' from left to right therefore can be slow.  If the same logical line is
    ' searched multiple times for different tokens, we want to ensure that the
    ' ToolParseLogicalLine procedure is only called once instead of calling it
    ' everytime a token is searched for.

    ' Get length of token
    Dim lngTokenLength As Long
    lngTokenLength = Len(strToken)

    Do

        ' Find first/next token position
        Dim lngTokenPosition As Long
        If fSearchForward Then
            lngTokenPosition = InStr(lngSearchStart, strSearchLogicalLine, strToken, vbBinaryCompare)
        Else
            lngTokenPosition = InStrRev(strSearchLogicalLine, strToken, lngSearchStart, vbBinaryCompare)
        End If

        ' Verify token position found
        If lngTokenPosition <> 0 Then
            ' Verify token delimited if token type needs to be delimited
            ' Note: We verify token is delimited to verify we did not find a parital match.
            If ToolTokenIsDelimitedIfNeeded(strSearchLogicalLine, lngTokenPosition, lngTokenLength, tttTokenType) Then
                ' Token found, return position and exit loop
                ToolGetTokenPosition = lngTokenPosition
                Exit Do
            Else
                ' Update search start to 1 character after last found token
                If fSearchForward Then
                    lngSearchStart = lngTokenPosition + lngTokenLength
                Else
                    lngSearchStart = lngTokenPosition - 1
                End If
            End If
        Else
            ' Token not found, exit loop
            Exit Do
        End If

    Loop

End Function

#If False Then

Public Sub ToolListVBComponents()

    #If False Then

        Dim lngComponentIndex As Long
        For lngComponentIndex = 1 To Application.VBE.ActiveVBProject.VBComponents.Count - 1
            Debug.Print Application.VBE.ActiveVBProject.VBComponents(lngComponentIndex).Name
        Next

    #Else

        Dim objVBComponent As VBIDE.VBComponent
        For Each objVBComponent In Application.VBE.ActiveVBProject.VBComponents
            ' objVBComponent.Type
            ' Module = 1 ' vbext_ct_StdModule
            ' Class = 2 ' vbext_ct_ClassModule
            ' Form = 100 ' vbext_ct_Document
            ' Report = 100 ' vbext_ct_Document
            Debug.Print objVBComponent.Name & " Type=" & objVBComponent.Type = vbext_ct_Document
            'ToolListCodeModuleProcedures objVBComponent.CodeModule
        Next

    #End If

End Sub

Public Sub ToolListCodeModuleProcedures(ByRef objCodeModule As VBIDE.CodeModule)

    With objCodeModule

        ' Cycle through all non declaration lines of code
        Dim lngNonDeclarationLine As Long
        For lngNonDeclarationLine = .CountOfDeclarationLines + 1 To .CountOfLines

            ' Get procedure name for current line of code
            ' Notes:
            ' vbext_ProcKind
            ' vbext_pk_Get  Specifies a procedure that returns the value of a property.
            ' vbext_pk_Let  Specifies a procedure that assigns a value to a property.
            ' vbext_pk_Set  Specifies a procedure that sets a reference to an object.
            ' vbext_pk_Proc Specifies all procedures other than property procedures.
            Dim strProcedureName As String
            Dim lngProcedureType As vbext_ProcKind
            strProcedureName = .ProcOfLine(lngNonDeclarationLine, lngProcedureType)

            ' Verify we found a procedure name
            If LenB(strProcedureName) <> 0 Then

                Debug.Print "Line " & lngNonDeclarationLine, strProcedureName
                Debug.Print "Line " & lngNonDeclarationLine, ToolGetLogicalLine(objCodeModule, .ProcBodyLine(strProcedureName, lngProcedureType))

                ' Skip remaining lines of code in procedure
                lngNonDeclarationLine = lngNonDeclarationLine + .ProcCountLines(strProcedureName, lngProcedureType) - 1

            End If

        Next

    End With

End Sub

Public Sub ToolFindValidIdentifierAndKeywordDelimiters()

Dim i As Long
For i = 0 To 255
    Debug.Print "Rem" & Chr(i) & "' " & i
Next

' Above code attemps to right delimit Rem keyword will all characters that
' have a character code from 0 to 255. The Rem keyword was chosen because
' it starts a comment therefore any character after Rem keyword should
' be valid as long as it is delimits the Rem keyword.
' Results of test indicate that:
' 1. No character with a character code > 127 can be used as a delimiter.
' 2. All characters with a character code =< 127 can be used as a delimiter
'    except numbers, letters and the underscore character.
' This makes since considering the following identifier/keyword naming rules:
' 1. Must begin with a letter.
' 2. Must contain only letters, numbers, and underscores (no other symbols).
' 3. Cannot be more than 255 characters long.

' Valid delimiter characters were found by attempting to right delimit the
' Rem keyword with all characters that have a character code from 0 to 255.
' The Rem keyword was chosen because it starts a comment therefore any
' character after Rem keyword should be valid as long as it delimits the
' Rem keyword.  Test results indicated:
' 1. No characters that have a character code > 127 can be used as a delimiter.
' 2. All characters that have a character code =< 127 can be used as a delimiter
'    except numbers, letters and the underscore character.
' This is consistent with the following identifier and keyword naming rules:
' 1. Must begin with a letter.
' 2. Must contain only letters, numbers, and underscores (no other symbols).
' 3. Cannot be more than 255 characters long.
' The fact that a character can be used to delimit an identifer or keyword
' token does not mean that character is valid.  It only means that the
' character can be used to reliable indicate the start and end of an
' identifier or keyword token.

#If False Then
Rem ' 0
Rem' 1
Rem' 2
Rem' 3
Rem' 4
Rem' 5
Rem' 6
Rem' 7
Rem' 8
Rem ' 9
Rem
' 10
Rem' 11
Rem' 12
Rem
' 13
Rem' 14
Rem' 15
Rem' 16
Rem' 17
Rem' 18
Rem' 19
Rem' 20
Rem' 21
Rem' 22
Rem' 23
Rem' 24
Rem' 25
Rem' 26
Rem' 27
Rem' 28
Rem' 29
Rem' 30
Rem' 31
Rem ' 32
Rem' 33
Rem"' 34
Rem' 35
Rem' 36
Rem' 37
Rem' 38
Rem'' 39
Rem(' 40
Rem)' 41
Rem*' 42
Rem+' 43
Rem,' 44
Rem-' 45
Rem.' 46
Rem/' 47
Rem0 ' 48
Rem1 ' 49
Rem2 ' 50
Rem3 ' 51
Rem4 ' 52
Rem5 ' 53
Rem6 ' 54
Rem7 ' 55
Rem8 ' 56
Rem9 ' 57
Rem:' 58
Rem;' 59
Rem<' 60
Rem=' 61
Rem>' 62
Rem?' 63
Rem' 64
Rema ' 65
Remb ' 66
Remc ' 67
Remd ' 68
Reme ' 69
Remf ' 70
Remg ' 71
Remh ' 72
Remi ' 73
Remj ' 74
Remk ' 75
Reml ' 76
Remm ' 77
Remn ' 78
Remo ' 79
Remp ' 80
Remq ' 81
Remr ' 82
Rems ' 83
Remt ' 84
Remu ' 85
Remv ' 86
Remw ' 87
Remx ' 88
Remy ' 89
Remz ' 90
Rem[' 91
Rem\' 92
Rem]' 93
Rem^' 94
Rem_ ' 95
Rem`' 96
Rema ' 97
Remb ' 98
Remc ' 99
Remd ' 100
Reme ' 101
Remf ' 102
Remg ' 103
Remh ' 104
Remi ' 105
Remj ' 106
Remk ' 107
Reml ' 108
Remm ' 109
Remn ' 110
Remo ' 111
Remp ' 112
Remq ' 113
Remr ' 114
Rems ' 115
Remt ' 116
Remu ' 117
Remv ' 118
Remw ' 119
Remx ' 120
Remy ' 121
Remz ' 122
Rem{' 123
Rem|' 124
Rem}' 125
Rem~' 126
Rem' 127
RemÄ ' 128
RemÅ ' 129
RemÇ ' 130
RemÉ ' 131
RemÑ ' 132
RemÖ ' 133
RemÜ ' 134
Remá ' 135
Remà ' 136
Remâ ' 137
Remö ' 138
Remã ' 139
Remú ' 140
Remç ' 141
Remû ' 142
Remè ' 143
Remê ' 144
Remë ' 145
Remí ' 146
Remì ' 147
Remî ' 148
Remï ' 149
Remñ ' 150
Remó ' 151
Remò ' 152
Remô ' 153
Remö ' 154
Remõ ' 155
Remú ' 156
Remù ' 157
Remû ' 158
Remˇ ' 159
Rem† ' 160
Rem° ' 161
Rem¢ ' 162
Rem£ ' 163
Rem§ ' 164
Rem• ' 165
Rem¶ ' 166
Remß ' 167
Rem® ' 168
Rem© ' 169
Rem™ ' 170
Rem´ ' 171
Rem¨ ' 172
Rem≠ ' 173
RemÆ ' 174
RemØ ' 175
Rem∞ ' 176
Rem± ' 177
Rem2 ' 178
Rem3 ' 179
Rem¥ ' 180
Remµ ' 181
Rem∂ ' 182
Rem∑ ' 183
Rem∏ ' 184
Rem1 ' 185
Rem∫ ' 186
Remª ' 187
Remº ' 188
RemΩ ' 189
Remæ ' 190
Remø ' 191
Rem‡ ' 192
Rem· ' 193
Rem‚ ' 194
Rem„ ' 195
Rem‰ ' 196
RemÂ ' 197
RemÊ ' 198
RemÁ ' 199
RemË ' 200
RemÈ ' 201
RemÍ ' 202
RemÎ ' 203
RemÏ ' 204
RemÌ ' 205
RemÓ ' 206
RemÔ ' 207
Rem ' 208
RemÒ ' 209
RemÚ ' 210
RemÛ ' 211
RemÙ ' 212
Remı ' 213
Remˆ ' 214
Rem◊ ' 215
Rem¯ ' 216
Rem˘ ' 217
Rem˙ ' 218
Rem˚ ' 219
Rem¸ ' 220
Rem˝ ' 221
Rem˛ ' 222
Remﬂ ' 223
Rem‡ ' 224
Rem· ' 225
Rem‚ ' 226
Rem„ ' 227
Rem‰ ' 228
RemÂ ' 229
RemÊ ' 230
RemÁ ' 231
RemË ' 232
RemÈ ' 233
RemÍ ' 234
RemÎ ' 235
RemÏ ' 236
RemÌ ' 237
RemÓ ' 238
RemÔ ' 239
Rem ' 240
RemÒ ' 241
RemÚ ' 242
RemÛ ' 243
RemÙ ' 244
Remı ' 245
Remˆ ' 246
Rem˜ ' 247
Rem¯ ' 248
Rem˘ ' 249
Rem˙ ' 250
Rem˚ ' 251
Rem¸ ' 252
Rem˝ ' 253
Rem˛ ' 254
Remˇ ' 255
#End If

End Sub

Public Function ToolTestToolGetBlankedLogicalLine()

    ' Test Results:

    ' Create copy of string inside function and return copy of string
    ' 3632 to 3652

    ' Create copy of string and pass to procedure instead of function
    ' 3539 to 3551

    ' Use function name inside function instead of creating a copy of string inside function
    ' 3624 to 3637

    Dim strSearchLogicalLine As String
    Dim lngCounter As Long

    ToolStartTimer
    For lngCounter = 1 To 100000
        strSearchLogicalLine = ToolGetBlankedLogicalLine("frm1.txtControl1.Value10A111 a #DateLiteral# b ""StringLiteral"" c 1.1# d &HFF e &H77 f ' Comment")
    Next
    Debug.Print ToolStopTimer

End Function

#End If




