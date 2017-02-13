Attribute VB_Name = "basConvertDocToMarkdown"
Option Explicit
'Authored 2015-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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

Public Sub ConvertWordDocumentToMarkdownText()
Dim doc As Document
Set doc = ActiveDocument
Dim stl As Style
Dim sty As Range
Dim par As Paragraph

    For Each sty In doc.StoryRanges
        Dim h As Hyperlink
        Dim rng As Range
        For Each h In sty.Hyperlinks
            If Len(h.Address) > 0 Then
                If Len(h.TextToDisplay) > 0 Then
                    Set rng = h.Range
                    rng.StartOf
                    rng.SetRange rng.Start, rng.Start
                    rng.Text = "["
                    Set rng = h.Range
                    rng.EndOf
                    rng.SetRange rng.End, rng.End
                    rng.Text = "](" & h.Address & ")"
                    If Right(h.TextToDisplay, 1) = " " Then
                        rng.EndOf
                        rng.SetRange rng.End, rng.End
                        rng.Text = rng.Text & " "
                    End If
                Else
                    Dim image As Shape
                    Set image = h.Shape
                    Set rng = image.Anchor
                    rng.EndOf
                    rng.SetRange rng.End, rng.End
                    Dim strHtml As String
                    strHtml = "<img src=""" & h.Address & """ alt=""" & image.Name & """>"
                    rng.Text = strHtml
                End If
            End If
        Next
    Next
    Dim strContent As String
    For Each par In doc.Paragraphs
        Set stl = par.Style
        Select Case LCase(stl.NameLocal)
            Case Is = "heading 1"
                par.Range.Text = "#" & par.Range.Text
            Case Is = "heading 2"
                par.Range.Text = "##" & par.Range.Text
            Case Is = "heading 3"
                par.Range.Text = "###" & par.Range.Text
            Case Is = "heading 4"
                par.Range.Text = "####" & par.Range.Text
        End Select
        strContent = strContent & vbCrLf & par.Range.Text
    Next
    Dim strTempFilePath As String
    strTempFilePath = Environ("TEMP") & Format(Now(), "yymmddhhss") & Right(Timer, 2) & ".md"
    SaveStringToFile strTempFilePath, strContent
    OpenFileWithExplorer strTempFilePath, False
    doc.Close False
End Sub

Private Sub SaveStringToFile(ByRef strFilePath As String, ByRef strString As String)
On Error GoTo HandleError

Dim intFileNumber As Long
Dim abyteByteArray() As Byte

    ' Delete existing file if needed
    If LenB(Dir(strFilePath)) <> 0 Then _
        Kill strFilePath

    ' Get free file number
    intFileNumber = FreeFile
    ' Open file for binary write
    Open strFilePath For Binary Access Write As intFileNumber
    ' Convert string to byte array
    ' Note: Must save string as byte array or Put function
    ' will convert string from unicode to ANSI.
    ' Empty string will NOT cause error.
    abyteByteArray() = strString
    ' Save data to file
    ' Note: Unallocated array will NOT cause error.
    Put intFileNumber, 1, abyteByteArray()
    ' Close file
    Close intFileNumber

ExitHere:
    Exit Sub

HandleError:
    ' Close file if needed
    ' Note: Below line of code will not raise an error even if no file is open
    Close intFileNumber
    Select Case Err.Number
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select

End Sub


Private Sub OpenFileWithExplorer(ByRef strFilePath As String, Optional ByRef fReadOnly As Boolean = True)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    Set wshShell = Nothing

End Sub
