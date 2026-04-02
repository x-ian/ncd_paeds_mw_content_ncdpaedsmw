Sub SplitSaveChapters_WordPDF_WithTOC()
    Dim doc As Document
    Dim newDoc As Document
    Dim rng As Range
    Dim searchRange As Range
    Dim title As String
    Dim masterName As String
    Dim folderPath As String
    Dim chapterNum As Integer
    Dim endPos As Long
    Dim baseFileName As String
    Dim wordPath As String
    Dim pdfPath As String
    Dim formattedNum As String
    
    Set doc = ActiveDocument
    
    ' Check if the document has been saved
    If doc.Path = "" Then
        MsgBox "Please save your master document specifically to a folder first.", vbCritical
        Exit Sub
    End If
    
    ' Get folder path and master filename
    folderPath = doc.Path & "\"
    If InStrRev(doc.Name, ".") > 0 Then
        masterName = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
    Else
        masterName = doc.Name
    End If
    
    Set searchRange = doc.Range
    searchRange.Collapse Direction:=wdCollapseStart
    
    chapterNum = 0
    Application.ScreenUpdating = False
    
    With searchRange.Find
        .Format = True
        .Style = doc.Styles("Heading 1")
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            chapterNum = chapterNum + 1
            formattedNum = Format(chapterNum, "00")
            
            ' 1. Get Title
            title = searchRange.Paragraphs(1).Range.Text
            
            ' 2. Clean Title
            title = Replace(title, vbCr, "")
            title = Replace(title, vbLf, "")
            title = Replace(title, ":", "-")
            title = Replace(title, "\", "-")
            title = Replace(title, "/", "-")
            title = Replace(title, "*", "")
            title = Replace(title, "?", "")
            title = Replace(title, """", "")
            title = Replace(title, "<", "")
            title = Replace(title, ">", "")
            title = Replace(title, "|", "")
            title = Trim(title)
            
            ' 3. Find End Position
            Dim startPos As Long
            startPos = searchRange.Start
            
            Dim tempRange As Range
            Set tempRange = searchRange.Duplicate
            tempRange.Collapse Direction:=wdCollapseEnd
            
            With tempRange.Find
                .Format = True
                .Style = doc.Styles("Heading 1")
                .Forward = True
                .Wrap = wdFindStop
                If .Execute Then
                    endPos = tempRange.Start
                Else
                    endPos = doc.Range.End
                End If
            End With
            
            ' 4. Copy Content
            Set rng = doc.Range(startPos, endPos)
            rng.Copy
            
            ' 5. Create New Doc (Clone Template)
            Set newDoc = Documents.Add(Template:=doc.FullName, Visible:=False)
            newDoc.Content.Delete
            newDoc.Content.Paste
            
            ' 6. Construct Base Filename
            baseFileName = masterName & "-" & title
            baseFileName = Replace(baseFileName, " ", "")
            
            ' 7. Save as Word (.docx)
            wordPath = folderPath & baseFileName & ".docx"
            newDoc.SaveAs2 FileName:=wordPath, FileFormat:=wdFormatXMLDocument
            
            FormatDocument newDoc
            
            ' 8. Save as PDF (.pdf) - WITH BOOKMARKS ENABLED
            pdfPath = folderPath & baseFileName & ".pdf"
            newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, _
                ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, _
                OptimizeFor:=wdExportOptimizeForPrint, _
                Range:=wdExportAllDocument, _
                Item:=wdExportDocumentContent, _
                IncludeDocProps:=True, _
                KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateHeadingBookmarks, _
                DocStructureTags:=True, _
                BitmapMissingFonts:=True, _
                UseISO19005_1:=False
            
            newDoc.Save
            newDoc.Close
            
            ' 9. Loop
            searchRange.Start = endPos
            searchRange.End = doc.Range.End
            If searchRange.Start >= doc.Range.End Then Exit Do
        Loop
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Success! Word docs and navigable PDFs created."
End Sub


Sub FormatDocument(doc As Document)
    
'    Dim doc As Document
'    Set doc = ActiveDocument
    
    ' --- Step 1: Move cursor above the colored horizontal line after "Chapter X" heading ---
    ' Find the first Heading 1 paragraph, then position cursor on the empty line after it
    Dim para As Paragraph
    Dim foundHeading As Boolean
    foundHeading = False
    
    For Each para In doc.Paragraphs
        If para.Style = doc.Styles(wdStyleHeading1) Then
            foundHeading = True
        ElseIf foundHeading Then
            ' This is the paragraph right after Heading 1
            ' Move selection to the start of this paragraph (the empty line above the horizontal line)
            para.Range.Select
            Selection.Collapse Direction:=wdCollapseStart
            Exit For
        End If
    Next para
    
    If Not foundHeading Then
        MsgBox "No Heading 1 found.", vbExclamation
        Exit Sub
    End If
    
    ' --- Step 2: Insert Table of Contents, excluding Heading 1 ---
    ' This inserts a TOC using Heading 2 and Heading 3 only (levels 2-3)
    Dim tocRange As Range
    Set tocRange = Selection.Range
    
    doc.TablesOfContents.Add _
        Range:=tocRange, _
        UseHeadingStyles:=True, _
        UpperHeadingLevel:=2, _
        LowerHeadingLevel:=3, _
        UseFields:=False, _
        TableID:="", _
        RightAlignPageNumbers:=True, _
        IncludePageNumbers:=True, _
        AddedStyles:="", _
        UseHyperlinks:=True, _
        HidePageNumbersInWeb:=True, _
        UseOutlineLevels:=True
    
    ' --- Step 3: Insert page numbers ---
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.Footers(wdHeaderFooterPrimary)
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, _
                            FirstPage:=True
        End With
    Next sec
    
End Sub

Sub FormatAndExportDocument()
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' --- Step 1: Move cursor above the colored horizontal line after "Chapter X" heading ---
    ' Find the first Heading 1 paragraph, then position cursor on the empty line after it
    Dim para As Paragraph
    Dim foundHeading As Boolean
    foundHeading = False
    
    For Each para In doc.Paragraphs
        If para.Style = doc.Styles(wdStyleHeading1) Then
            foundHeading = True
        ElseIf foundHeading Then
            ' This is the paragraph right after Heading 1
            ' Move selection to the start of this paragraph (the empty line above the horizontal line)
            para.Range.Select
            Selection.Collapse Direction:=wdCollapseStart
            Exit For
        End If
    Next para
    
    If Not foundHeading Then
        MsgBox "No Heading 1 found.", vbExclamation
        Exit Sub
    End If
    
    ' --- Step 2: Insert Table of Contents, excluding Heading 1 ---
    ' This inserts a TOC using Heading 2 and Heading 3 only (levels 2-3)
    Dim tocRange As Range
    Set tocRange = Selection.Range
    
    doc.TablesOfContents.Add _
        Range:=tocRange, _
        UseHeadingStyles:=True, _
        UpperHeadingLevel:=2, _
        LowerHeadingLevel:=3, _
        UseFields:=False, _
        TableID:="", _
        RightAlignPageNumbers:=True, _
        IncludePageNumbers:=True, _
        AddedStyles:="", _
        UseHyperlinks:=True, _
        HidePageNumbersInWeb:=True, _
        UseOutlineLevels:=True
    
    ' --- Step 3: Insert page numbers ---
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.Footers(wdHeaderFooterPrimary)
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, _
                            FirstPage:=True
        End With
    Next sec
    
    ' --- Step 4: Save As PDF (optimized for electronic publishing) ---
    Dim pdfPath As String
    pdfPath = Replace(doc.FullName, ".docm", ".pdf")
    pdfPath = Replace(pdfPath, ".docx", ".pdf")
    pdfPath = Replace(pdfPath, ".doc", ".pdf")
    
    doc.ExportAsFixedFormat _
        OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForOnScreen, _
        Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False
    
    ' --- Step 5: Save the document ---
    doc.Save
    
    MsgBox "Done. PDF saved to:" & vbCrLf & pdfPath, vbInformation
    
End Sub

