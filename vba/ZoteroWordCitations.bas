Attribute VB_Name = "ZoteroWordCitations"
Option Explicit

' Change this to a full python.exe path if Word cannot find the Python used to
' install zotero-word-citations, for example:
' C:\Users\your-name\AppData\Local\Programs\Python\Python312\python.exe
Private Const PYTHON_COMMAND As String = "py"

Public Sub ZoteroCiteDOIsInActiveDocument()
    Dim sourceDocument As Document
    Dim sourcePath As String
    Dim outputPath As String
    Dim commandLine As String
    Dim exitCode As Long
    Dim shell As Object

    On Error GoTo Failed

    If Documents.Count = 0 Then
        MsgBox "Open a Word document first.", vbInformation, "Zotero Word Citations"
        Exit Sub
    End If

    Set sourceDocument = ActiveDocument
    If Len(sourceDocument.Path) = 0 Then
        MsgBox "Save the document as a .docx file before running this command.", _
               vbInformation, "Zotero Word Citations"
        Exit Sub
    End If

    sourcePath = sourceDocument.FullName
    If LCase$(Right$(sourcePath, 5)) <> ".docx" Then
        MsgBox "The active document must be saved in .docx format.", _
               vbExclamation, "Zotero Word Citations"
        Exit Sub
    End If

    If Not sourceDocument.Saved Then sourceDocument.Save
    outputPath = Left$(sourcePath, Len(sourcePath) - 5) & "_zotero_cited.docx"

    If IsDocumentOpen(outputPath) Then
        MsgBox "Close the existing output document before running again:" & _
               vbCrLf & outputPath, vbExclamation, "Zotero Word Citations"
        Exit Sub
    End If

    commandLine = QuoteArgument(PYTHON_COMMAND) & _
                  " -m zotero_word_citations " & QuoteArgument(sourcePath) & _
                  " --force"

    Application.StatusBar = "Converting DOI placeholders to Zotero citations..."
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(commandLine, 1, True)
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "The converter returned error code " & CStr(exitCode) & "." & _
               vbCrLf & "Review the terminal window and confirm that Zotero is running.", _
               vbCritical, "Zotero Word Citations"
        Exit Sub
    End If

    If Len(Dir$(outputPath)) = 0 Then
        MsgBox "No output document was created. No uncited DOI placeholders may remain." & _
               vbCrLf & "See dois.txt beside the source document for details.", _
               vbInformation, "Zotero Word Citations"
        Exit Sub
    End If

    Documents.Open FileName:=outputPath, AddToRecentFiles:=True
    MsgBox "Created and opened:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
           "Use Zotero > Refresh to finalize the citation style and numbering.", _
           vbInformation, "Zotero Word Citations"
    Exit Sub

Failed:
    Application.StatusBar = False
    MsgBox "The conversion could not start: " & Err.Description, _
           vbCritical, "Zotero Word Citations"
End Sub

Private Function QuoteArgument(ByVal value As String) As String
    QuoteArgument = Chr$(34) & value & Chr$(34)
End Function

Private Function IsDocumentOpen(ByVal candidatePath As String) As Boolean
    Dim document As Document

    For Each document In Application.Documents
        If StrComp(document.FullName, candidatePath, vbTextCompare) = 0 Then
            IsDocumentOpen = True
            Exit Function
        End If
    Next document
End Function
