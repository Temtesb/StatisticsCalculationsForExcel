Attribute VB_Name = "build_ImportExportBas"
Option Explicit
'Authored 2014-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
    'Public Domain in the United States of America,
     'any international rights are relinquished under CC0 1.0 <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
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

#If True Then 'can only execute on a macro trusted platform

Public Sub ToolExportAllModualsAsBas()
Dim strComponentName As String, strCurrent_FilePath As String
Dim fso As Object, objFile As Object, objFolder As Object
Dim intCurrentComponent As Integer, intVbCompontentsCount As Integer
'Force late binding on applicaton so we can compile and only at runtime do those lines of code execute that are appropriate
    Set fso = CreateObject("Scripting.FileSystemObject")
    strCurrent_FilePath = GetRelativePathViaParent()
    'Folder structure is:
        'Excel only code: /Excel/src
        'Access only code: /Access/src
        'Generic code for all (name begins with 'VB_'): /Generic VB/src
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Ensure the Export Directories exists, build them if needed
    intVbCompontentsCount = Application.VBE.ActiveVBProject.VBComponents.Count
    Dim strDestinationPath As String
    For intCurrentComponent = 1 To intVbCompontentsCount
        strComponentName = Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent).Name
        If Left(strComponentName, 5) <> "Form_" And Left(strComponentName, 7) <> "Report_" Then
            Select Case True
                Case LCase(Left(strComponentName, 3)) = "vb_"
                    strDestinationPath = GetRelativePathViaParent("..\..\General VB\src")
                Case LCase(Left(strComponentName, 4)) = "bas_"
                    strDestinationPath = GetRelativePathViaParent("..\src")
                Case LCase(Left(strComponentName, 6)) = "build_"
                    strDestinationPath = GetRelativePathViaParent("..\..\build")
                Case Else
                    GoTo NextComponent
            End Select
            If Not fso.FolderExists(strDestinationPath) Then
                BuildDir strDestinationPath
            End If
            Dim objTest As Object
            Set objTest = Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent)
            Select Case Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent).Type
                Case vbext_ct_ClassModule
                    Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent).Export strDestinationPath & "\" & strComponentName & ".cl"
                Case vbext_ct_StdModule
                    Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent).Export strDestinationPath & "\" & strComponentName & ".bas"
                Case Else
                    Application.VBE.ActiveVBProject.VBComponents.Item(intCurrentComponent).Export strDestinationPath & "\" & strComponentName
            End Select
        End If
NextComponent:
    Next intCurrentComponent
    Debug.Print "Export complete"
End Sub

Public Sub ToolImportModules()
    'This tool can be loaded to a file in the main root folder, build directory, or the excel or access folder
    mImportVbComponent GetRelativePathViaParent("..\..\General VB\src")
    mImportVbComponent GetRelativePathViaParent("..\General VB\src")
    mImportVbComponent GetRelativePathViaParent("General VB\src")
    mImportVbComponent GetRelativePathViaParent("..\src")
    mImportVbComponent GetRelativePathViaParent("src")
    mImportVbComponent GetRelativePathViaParent("..\..\Build")
    mImportVbComponent GetRelativePathViaParent("..\Build")
    mImportVbComponent GetRelativePathViaParent("Build")
    Dim oThisApplication As Object
    Set oThisApplication = Application
    Select Case True
        Case InStrRev(oThisApplication.Name, "Excel") > 0
                mImportVbComponent GetRelativePathViaParent("Excel\src")
                mImportVbComponent GetRelativePathViaParent("Excel\Build")
        Case InStrRev(oThisApplication.Name, "Access") > 0
                mImportVbComponent GetRelativePathViaParent("Access\src")
                mImportVbComponent GetRelativePathViaParent("Access\Build")
        'Case InStrRev(oThisApplication.Name, "Word") > 0
    End Select
End Sub

Private Sub mImportVbComponent(strFolderSource)
    Dim fil As Object, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(strFolderSource) Then
        For Each fil In fso.GetFolder(strFolderSource).Files
            Dim strExtension As String: strExtension = LCase(mGetFileExtension(fil.Name))
            If strExtension = "cl" Or strExtension = "bas" Or Len(strExtension) = 0 Then
                If VbComponentExits(fil.Name) Then
                    If MsgBox("Do you want to replace the existing Visual Basic Component with the file:" & fil.Name & " - last modified:" & fil.DateLastModified & "?", vbYesNo + vbQuestion, "VBA Statistics Import Conflict") = vbYes Then
                        Application.VBE.ActiveVBProject.VBComponents.Import (fil.Path)
                    End If
                Else
                    Application.VBE.ActiveVBProject.VBComponents.Import (fil.Path)
                End If
            End If
        Next
    End If
End Sub
Private Function mGetFileExtension(strFileName)
    mGetFileExtension = Right(strFileName, Len(strFileName) - InStrRev(strFileName, "."))
End Function

Private Function VbComponentExits(strFileName As String) As Boolean
    On Error Resume Next
    Dim strVBCName As String
    strVBCName = Application.VBE.ActiveVBProject.VBComponents(Left(strFileName, InStrRev(strFileName, ".") - 1)).Name
    VbComponentExits = Err.Number = 0
End Function
#End If

