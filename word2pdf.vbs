Option Explicit

Dim wordApp, inputFolder, pdfFolder, fso, file, doc, pdfFile
Dim pdfFiles, mergerFile, pdfMerger

' Define the input and output folders
inputFolder = "D:\backup\guan\output1" ' Change this path
pdfFolder = "D:\backup\guan\output2" ' Change this path

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Create the output PDF folder if it doesn't exist
If Not fso.FolderExists(pdfFolder) Then
    fso.CreateFolder(pdfFolder)
End If

' Initialize Word application
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

' Convert .docx files to .pdf
For Each file In fso.GetFolder(inputFolder).Files
    If LCase(fso.GetExtensionName(file.Name)) = "docx" Then
        Set doc = wordApp.Documents.Open(file.Path)
        pdfFile = fso.BuildPath(pdfFolder, fso.GetBaseName(file.Name) & ".pdf")
        doc.SaveAs pdfFile, 17 ' 17 is the format code for PDF
        doc.Close False
    End If
Next

' Quit Word application
wordApp.Quit

' Merge PDFs using GhostScript
Set pdfFiles = CreateObject("Scripting.Dictionary")
For Each file In fso.GetFolder(pdfFolder).Files
    If LCase(fso.GetExtensionName(file.Name)) = "pdf" Then
        pdfFiles.Add pdfFiles.Count, file.Path
    End If
Next

' Prepare the command for merging PDFs
mergerFile = fso.BuildPath(pdfFolder, "merged_output.pdf")
Dim gsCommand, gsPath
gsPath = "C:\Program Files\gs\gsX.X\bin\gswin64c.exe" ' Update with your Ghostscript path

gsCommand = """" & gsPath & """ -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile=""" & mergerFile & """ "
For Each pdfFile In pdfFiles.Keys
    gsCommand = gsCommand & """" & pdfFiles(pdfFile) & """ "
Next

' Execute the GhostScript command
CreateObject("WScript.Shell").Run gsCommand, 0, True

WScript.Echo "All .docx files have been converted to PDF and merged into '" & mergerFile & "'."
