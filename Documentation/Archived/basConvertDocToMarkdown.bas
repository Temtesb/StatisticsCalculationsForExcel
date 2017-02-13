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

Private Function HyperlinkHasAddress(h As Hyperlink) As Boolean
On Error Resume Next
    Dim lngAddressLength
    lngAddressLength = 0
    lngAddressLength = Len(h.Address)
    HyperlinkHasAddress = lngAddressLength > 0
End Function

Public Sub ConvertWordDocumentToMarkdownText()
Dim doc As Document
Set doc = ActiveDocument
Dim stl As Style
Dim sty As Range
Dim par As Paragraph
    Dim h As Hyperlink
    Dim rng As Range
    Dim shp As InlineShape
    For Each shp In doc.InlineShapes
        Set h = shp.Hyperlink
            
        If HyperlinkHasAddress(h) Then
            Set rng = shp.Range
            rng.EndOf
            rng.SetRange rng.End, rng.End
            Dim strHtml As String
            strHtml = "<img src=""" & Trim(h.Address) & """ alt=""" & Trim(shp.Title) & """>"
            rng.Text = strHtml
        End If
        Set h = Nothing
    Next
    
    For Each sty In doc.StoryRanges
        For Each h In sty.Hyperlinks
            Debug.Print h.Address
            If Len(h.Address) > 0 Then
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
            End If
            For Each par In sty.ListParagraphs
                par.Range.Text = "  *" & par.Range.Text 'Should check for the depth of the list item
            Next
            
        Next
    Next
    Dim strContent As String
    'Count Paragraphs
    Dim lngParagraph As Long
    Dim lngParagraphCount As Long
    Dim objParagraphs As Paragraphs
    Set objParagraphs = doc.Paragraphs
    lngParagraphCount = objParagraphs.Count
    For lngParagraph = 1 To lngParagraphCount
        Set par = objParagraphs(lngParagraph)
        Set stl = par.Style
        If Len(par.Range.Text) > 2 Then
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
        End If
    Next
'    For Each par In objParagraphs
'        strContent = strContent & Chr(10) & par.Range.Text
'    Next
    Dim strTempFilePath As String
    strTempFilePath = doc.Path & "\" & Left(doc.Name, InStrRev(doc.Name, ".") - 1) & ".md"
'    SaveStringToFile strTempFilePath, strContent
    doc.SaveAs2 strTempFilePath, WdSaveFormat.wdFormatEncodedText, AddToRecentFiles:=False, Encoding:=MsoEncoding.msoEncodingUTF8
    OpenFileWithExplorer strTempFilePath, False
End Sub

Private Function isShape(obj As Variant) As Boolean
    On Error Resume Next
    Dim objTest As Shape
    Set objTest = obj
    isShape = Err.Number = 0
End Function
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
